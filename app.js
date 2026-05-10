let allDecks = JSON.parse(localStorage.getItem('myFlashcardDecks')) || {};
let currentDeckName = null;
let currentCardIndex = -1;
let showingAnswer = false;
let tempCardsArray = []; // Guarda temporalmente las cartas del Excel antes de confirmar
let isReverseMode = JSON.parse(localStorage.getItem('flashcards_reverse')) || false;

// Registro del Service Worker para funcionamiento offline
if ('serviceWorker' in navigator) {
    navigator.serviceWorker.register('./sw.js');
}

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
    
    // Filtramos los índices pendientes
    const dueIndices = deck.map((c, i) => c.nextReview <= now ? i : null).filter(i => i !== null);
    
    // Actualizamos el contador con el número REAL de cartas que esperan
    document.getElementById('cards-left').innerText = dueIndices.length;

    if (dueIndices.length === 0) {
        alert("¡Felicidades! Has terminado por hoy.");
        goToHome();
        return;
    }

    // Prioridad: Si hay cartas con nextReview = 0 (marcadas para repetir ya),
    // intentamos que no sea la misma que acabamos de ver si hay otras opciones.
    let possibleIndices = dueIndices;
    if (dueIndices.length > 1 && currentCardIndex !== -1) {
        possibleIndices = dueIndices.filter(i => i !== currentCardIndex);
    }

    currentCardIndex = possibleIndices[Math.floor(Math.random() * possibleIndices.length)];
    showingAnswer = false;
    
    const cardData = deck[currentCardIndex];
    document.getElementById('card-front-text').innerText = isReverseMode ? cardData.a : cardData.q;
    document.getElementById('card-back-text').innerText = isReverseMode ? cardData.q : cardData.a;

    // Quitamos la rotación para mostrar la nueva pregunta
    document.getElementById('card').classList.remove('flipped');
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

// Nueva función para días personalizados
function setCustomSchedule() {
    const input = document.getElementById('custom-days');
    const days = parseInt(input.value);
    if (!isNaN(days) && days >= 0) {
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

