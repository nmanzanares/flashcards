// Registro del Service Worker (necesario para PWA)
if ('serviceWorker' in navigator) {
    navigator.serviceWorker.register('sw.js');
}

let allDecks = JSON.parse(localStorage.getItem('myFlashcardDecks')) || {};
let currentDeckName = null;
let currentCardIndex = -1;
let showingAnswer = false;

document.addEventListener('DOMContentLoaded', () => {
    // 1. Renderizar mazos al abrir
    renderDecks();

    // 2. Evento para el clic en la tarjeta
    const cardElement = document.getElementById('card');
    if (cardElement) {
        cardElement.addEventListener('click', () => {
            if (currentCardIndex === -1 || showingAnswer) return;
            showingAnswer = true;
            document.getElementById('card').innerText = allDecks[currentDeckName][currentCardIndex].a;
            document.getElementById('answer-controls').style.display = 'block';
        });
    }

    // 3. Evento para importar Excel
    const excelInput = document.getElementById('excel-input');
    if (excelInput) {
        excelInput.addEventListener('change', importExcel);
    }
});

function renderDecks() {
    const container = document.getElementById('decks-container');
    if (!container) return;
    container.innerHTML = '';
    const now = Date.now();

    Object.keys(allDecks).forEach(name => {
        const deck = allDecks[name];
        const dueCards = deck.filter(c => c.nextReview <= now);
        
        const div = document.createElement('div');
        div.style = "border: 1px solid #ccc; padding: 15px; margin: 10px 0; border-radius: 10px; background: #fafafa;";
        div.innerHTML = `
            <strong>${name}</strong><br>
            Total: ${deck.length} | Pendientes: ${dueCards.length}<br>
            <button onclick="startStudy('${name}')" ${dueCards.length === 0 ? 'disabled' : ''}>Estudiar</button>
            <button onclick="deleteDeck('${name}')" style="color: red;">Eliminar</button>
        `;
        container.appendChild(div);
    });
}

function startStudy(name) {
    currentDeckName = name;
    document.getElementById('current-deck-title').innerText = name;
    document.getElementById('setup-view').style.display = 'none';
    document.getElementById('study-view').style.display = 'block';
    showNextCard();
}

function showNextCard() {
    const now = Date.now();
    const deck = allDecks[currentDeckName];
    // Filtramos las cartas que toca estudiar
    const dueIndices = deck.map((c, i) => c.nextReview <= now ? i : null).filter(i => i !== null);
    
    if (dueIndices.length === 0) {
        alert("¡Mazo completado!");
        goToHome();
        return;
    }

    document.getElementById('cards-left').innerText = dueIndices.length;
    // Elegimos una al azar de las pendientes
    currentCardIndex = dueIndices[Math.floor(Math.random() * dueIndices.length)];
    showingAnswer = false;
    document.getElementById('card').innerText = deck[currentCardIndex].q;
    document.getElementById('answer-controls').style.display = 'none';
}

function setSchedule(days) {
    const msInDay = 24 * 60 * 60 * 1000;
    allDecks[currentDeckName][currentCardIndex].nextReview = Date.now() + (days * msInDay);
    
    localStorage.setItem('myFlashcardDecks', JSON.stringify(allDecks));
    showNextCard();
}

function goToHome() {
    document.getElementById('setup-view').style.display = 'block';
    document.getElementById('study-view').style.display = 'none';
    renderDecks();
}

function deleteDeck(name) {
    if (confirm(`¿Borrar ${name}?`)) {
        delete allDecks[name];
        localStorage.setItem('myFlashcardDecks', JSON.stringify(allDecks));
        renderDecks();
    }
}

function importExcel(e) {
    const name = document.getElementById('deck-name').value.trim();
    if (!name) return alert("Escribe un nombre para el mazo");

    const reader = new FileReader();
    reader.onload = function(event) {
        const data = new Uint8Array(event.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        const jsonData = XLSX.utils.sheet_to_json(workbook.Sheets[workbook.SheetNames]);

        allDecks[name] = jsonData.map(row => ({
            q: String(row.pregunta || row.Pregunta),
            a: String(row.respuesta || row.Respuesta),
            nextReview: 0
        }));

        localStorage.setItem('myFlashcardDecks', JSON.stringify(allDecks));
        document.getElementById('deck-name').value = '';
        renderDecks();
    };
    reader.readAsArrayBuffer(e.target.files[0]);
}


/*****
let deck = JSON.parse(localStorage.getItem('myFlashcards')) || [];
let currentCardIndex = -1;
let showingAnswer = false;

// 1. IMPORTACIÓN CON NUEVOS ATRIBUTOS
document.getElementById('excel-input').addEventListener('change', function(e) {
    const reader = new FileReader();
    reader.onload = function(event) {
        const data = new Uint8Array(event.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        const jsonData = XLSX.utils.sheet_to_json(workbook.Sheets[workbook.SheetNames[0]]);

        const newCards = jsonData.map(row => ({
            q: String(row.pregunta || row.Pregunta),
            a: String(row.respuesta || row.Respuesta),
            nextReview: Date.now(), // Disponible desde ya
            visible: true
        }));

        deck = [...deck, ...newCards];
        localStorage.setItem('myFlashcards', JSON.stringify(deck));
        updateInfo();
    };
    reader.readAsArrayBuffer(e.target.files[0]);
});

    

let allDecks = JSON.parse(localStorage.getItem('myFlashcardDecks')) || {};
let currentDeckName = null;

// Cargar mazos al inicio
function renderDecks() {
    const container = document.getElementById('decks-container');
    container.innerHTML = '';
    const now = Date.now();

    Object.keys(allDecks).forEach(name => {
        const deck = allDecks[name];
        const dueCount = deck.filter(c => c.visible && c.nextReview <= now).length;
        
        const div = document.createElement('div');
        div.className = 'deck-card';
        div.innerHTML = `
            <div style="border: 1px solid #ccc; padding: 10px; margin: 5px; border-radius: 8px;">
                <strong>${name}</strong><br>
                Total: ${deck.length} | Pendientes: <span style="color: red">${dueCount}</span><br>
                <button onclick="startStudy('${name}')" ${dueCount === 0 ? 'disabled' : ''}>Estudiar</button>
                <button onclick="deleteDeck('${name}')" style="background: #ff4444; color: white">Borrar</button>
            </div>
        `;
        container.appendChild(div);
    });
}

// Importar con nombre
document.getElementById('excel-input').addEventListener('change', function(e) {
    const name = document.getElementById('deck-name').value.trim();
    if (!name) return alert("Ponle un nombre al mazo primero");
    if (allDecks[name]) return alert("Ya existe un mazo con ese nombre");

    const reader = new FileReader();
    reader.onload = function(event) {
        const data = new Uint8Array(event.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        const jsonData = XLSX.utils.sheet_to_json(workbook.Sheets[workbook.SheetNames]);

        const newCards = jsonData.map(row => ({
            q: String(row.pregunta || row.Pregunta),
            a: String(row.respuesta || row.Respuesta),
            nextReview: 0, 
            visible: true
        }));

        allDecks[name] = newCards;
        localStorage.setItem('myFlashcardDecks', JSON.stringify(allDecks));
        document.getElementById('deck-name').value = '';
        renderDecks();
    };
    reader.readAsArrayBuffer(e.target.files);
});

function updateInfo() {
    const dueCards = deck.filter(c => c.visible && c.nextReview <= Date.now());
    document.getElementById('total-info').innerText = `Total: ${deck.length} | Pendientes hoy: ${dueCards.length}`;
}

// 2. CAMBIO DE PÁGINA


function startStudy(name) {
    currentDeckName = name;
    //const dueCards = deck.filter(c => c.visible && c.nextReview <= Date.now()); //Este 'deck' dara problema
    //if (dueCards.length === 0) return alert("¡No hay cartas pendientes para hoy!");
    document.getElementById('current-deck-title').innerText = name;
    document.getElementById('setup-view').style.display = 'none';
    document.getElementById('study-view').style.display = 'block';
    showNextCard();
}

function goToHome() {
    currentDeckName = null;
    document.getElementById('setup-view').style.display = 'block';
    document.getElementById('study-view').style.display = 'none';
    renderDecks();
}

function deleteDeck(name) {
    if (confirm(`¿Borrar el mazo ${name}?`)) {
        delete allDecks[name];
        localStorage.setItem('myFlashcardDecks', JSON.stringify(allDecks));
        renderDecks();
    }
}

// 3. LÓGICA DE ESTUDIO
function showNextCard() {
    const dueCards = allDecks[currentDeckName].filter(c => c.visible && c.nextReview <= Date.now());
    document.getElementById('cards-left').innerText = dueCards.length;

    if (dueCards.length === 0) {
        alert("¡Has terminado por ahora!");
        location.reload();
        return;
    }

    // Buscamos el índice real en el mazo original
    currentCardIndex = allDecks[currentDeckName].indexOf(dueCards[0]);
    showingAnswer = false;
    document.getElementById('card').innerText = allDecks[currentDeckName][currentCardIndex].q;
    document.getElementById('answer-controls').style.display = 'none';
}

document.getElementById('card').addEventListener('click', () => {
    if (currentCardIndex === -1 || showingAnswer) return;
    showingAnswer = true;
    document.getElementById('card').innerText = allDecks[currentDeckName][currentCardIndex].a;
    document.getElementById('answer-controls').style.display = 'block';
});

// 4. CONFIGURAR REPETICIÓN (DÍAS)
function setSchedule(days) {
    const msInDay = 24 * 60 * 60 * 1000;
    
    if (days === 0) {
        // Se queda en la cola de hoy (le ponemos un timestamp muy viejo para que siga saliendo)
        allDecks[currentDeckName][currentCardIndex].nextReview = Date.now();
    } else {
        // Se programa para el futuro
        allDecks[currentDeckName][currentCardIndex].nextReview = Date.now() + (days * msInDay);
    }

    localStorage.setItem('myFlashcardDecks', JSON.stringify(allDecks));
    showNextCard();
}


********/
// Inicializar info al cargar
updateInfo();






/***********
// Carga las tarjetas guardadas o usa las de ejemplo si no hay nada
let deck = JSON.parse(localStorage.getItem('myFlashcards')) || [
    { q: "No hay flashcards", a: "Error en la carga del fichero" }
];

let currentCard = 0;
let showingAnswer = false;
const cardElement = document.getElementById('card');

function nextCard() {
    currentCard = Math.floor(Math.random() * deck.length);
    showingAnswer = false;
    cardElement.innerText = deck[currentCard].q;
}

cardElement.addEventListener('click', () => {
    if (!deck[currentCard]) return;
    showingAnswer = !showingAnswer;
    cardElement.innerText = showingAnswer ? deck[currentCard].a : deck[currentCard].q;
});

document.getElementById('excel-input').addEventListener('change', function(e) {
    const file = e.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = function(event) {
        const data = new Uint8Array(event.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        const jsonData = XLSX.utils.sheet_to_json(sheet);

        console.log("Datos brutos del Excel:", jsonData); // Ver en consola del PC

        const newCards = jsonData.map(row => {
            // Esto busca la columna sin importar si es Mayúscula o Minúscula
            const q = row.pregunta || row.Pregunta || row.Question;
            const a = row.respuesta || row.Respuesta || row.Answer;
            return (q && a) ? { q: String(q).trim(), a: String(a).trim() } : null;
        }).filter(card => card !== null);

        if (newCards.length > 0) {
            // 1. Unir con las tarjetas que ya tenías
            const currentDeck = JSON.parse(localStorage.getItem('myFlashcards')) || [];
            const updatedDeck = [...currentDeck, ...newCards];
            
            // 2. Guardar permanentemente
            localStorage.setItem('myFlashcards', JSON.stringify(updatedDeck));
            
            // 3. Actualizar la app para que use las nuevas tarjetas
            location.reload(); 
            alert(`¡Éxito! Se han añadido ${newCards.length} tarjetas.`);
        } else {
            alert("No se encontraron datos válidos. Revisa los nombres de las columnas.");
        }
    };
    reader.readAsArrayBuffer(file);
});
****/


