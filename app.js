let allDecks = JSON.parse(localStorage.getItem('myFlashcardDecks')) || {};
let currentDeckName = null;
let currentCardIndex = -1;
let showingAnswer = false;
let tempCardsArray = []; // Guarda temporalmente las cartas del Excel antes de confirmar

// Registro del Service Worker para funcionamiento offline
if ('serviceWorker' in navigator) {
    navigator.serviceWorker.register('./sw.js');
}

document.addEventListener('DOMContentLoaded', () => {
    // Dibujar los mazos al cargar la página
    renderDecks();

    // Evento para voltear la tarjeta
    const cardElement = document.getElementById('card');
    if (cardElement) {
        cardElement.addEventListener('click', () => {
            if (currentCardIndex === -1 || showingAnswer) return;
            showingAnswer = true;
            document.getElementById('card').innerText = allDecks[currentDeckName][currentCardIndex].a;
            document.getElementById('answer-controls').style.display = 'block';
        });
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
        div.innerHTML = `
            <strong>${name}</strong><br>
            <span style="font-size: 0.9rem; color: #555;">
                Total cartas: ${deck.length} | 
                Para estudiar hoy: <b style="color: ${dueCardsCount > 0 ? '#ff4444' : 'green'};">${dueCardsCount}</b>
            </span>
            <div style="margin-top: 10px;">
                <button onclick="startStudy('${name}')" ${dueCardsCount === 0 ? 'disabled' : ''}>Estudiar</button>
                <button onclick="deleteDeck('${name}')" class="btn-danger">Eliminar</button>
            </div>
        `;
        container.appendChild(div);
    });
}

// Procesa el archivo Excel y lo guarda temporalmente en memoria
function handleExcelSelection(e) {
    const file = e.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = function(event) {
        try {
            const data = new Uint8Array(event.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            const jsonData = XLSX.utils.sheet_to_json(workbook.Sheets[workbook.SheetNames]);

            // Convertimos las filas del excel añadiendo por defecto nextReview = 0 (listas para estudiar)
            tempCardsArray = jsonData.map(row => ({
                q: String(row.pregunta || row.Pregunta || '').trim(),
                a: String(row.respuesta || row.Respuesta || '').trim(),
                nextReview: 0 
            })).filter(card => card.q !== '' && card.a !== ''); // Filtra filas vacías

            if(tempCardsArray.length === 0) {
                alert("Error: No se encontraron columnas válidas llamadas 'pregunta' y 'respuesta'.");
                e.target.value = '';
            } else {
                console.log("Archivo procesado. Listo para confirmar.", tempCardsArray);
            }
        } catch (err) {
            alert("Error al leer el archivo Excel.");
            e.target.value = '';
        }
    };
    reader.readAsArrayBuffer(file);
}

// Guarda definitivamente el mazo cuando el usuario pulsa "Confirmar y Añadir"
function confirmAndAddDeck() {
    const nameInput = document.getElementById('deck-name');
    const name = nameInput.value.trim();
    const fileInput = document.getElementById('excel-input');

    if (!name) return alert("Por favor, escribe un nombre para el mazo.");
    if (allDecks[name]) return alert("Ya existe un mazo con ese nombre. Elige otro.");
    if (tempCardsArray.length === 0) return alert("Por favor, selecciona primero un archivo Excel válido.");

    // Guardar en el objeto principal
    allDecks[name] = tempCardsArray;
    localStorage.setItem('myFlashcardDecks', JSON.stringify(allDecks));
    
    // Limpiar variables y formulario
    tempCardsArray = [];
    nameInput.value = '';
    fileInput.value = '';

    // Refrescar la vista
    renderDecks();
    alert(`¡Mazo "${name}" añadido correctamente!`);
}

// Inicia la pantalla de estudio de un mazo específico [7]
function startStudy(name) {
    currentDeckName = name;
    document.getElementById('current-deck-title').innerText = name;
    document.getElementById('setup-view').style.display = 'none';
    document.getElementById('study-view').style.display = 'block';
    showNextCard();
}

// Elige la siguiente carta pendiente al azar
function showNextCard() {
    const now = Date.now();
    const deck = allDecks[currentDeckName];
    
    // Buscamos los índices de las cartas que ya deben estudiarse
    const dueIndices = deck.map((c, i) => c.nextReview <= now ? i : null).filter(i => i !== null);
    
    document.getElementById('cards-left').innerText = dueIndices.length;

    // Si ya no quedan cartas pendientes, regresamos automáticamente a la pantalla de inicio [7]
    if (dueIndices.length === 0) {
        alert("¡Felicidades! Has terminado de estudiar este mazo por hoy.");
        goToHome();
        return;
    }

    // Elegir un índice al azar de los disponibles
    currentCardIndex = dueIndices[Math.floor(Math.random() * dueIndices.length)];
    showingAnswer = false;
    
    document.getElementById('card').innerText = deck[currentCardIndex].q;
    document.getElementById('answer-controls').style.display = 'none';
}

// Configura los días de retraso para la tarjeta actual
function setSchedule(days) {
    const msInDay = 24 * 60 * 60 * 1000;
    
    // Si elige 0 (repetir ya), nextReview se queda en 0 para que vuelva a salir en la misma sesión
    // Si elige días, sumamos los milisegundos correspondientes al tiempo actual
    allDecks[currentDeckName][currentCardIndex].nextReview = days === 0 ? 0 : Date.now() + (days * msInDay);
    
    localStorage.setItem('myFlashcardDecks', JSON.stringify(allDecks));
    showNextCard();
}

// Regresa a la pantalla principal y actualiza los contadores [7]
function goToHome() {
    currentDeckName = null;
    currentCardIndex = -1;
    document.getElementById('setup-view').style.display = 'block';
    document.getElementById('study-view').style.display = 'none';
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

