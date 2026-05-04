// Registro del Service Worker (necesario para PWA)
if ('serviceWorker' in navigator) {
    navigator.serviceWorker.register('sw.js');
}

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

function updateInfo() {
    const dueCards = deck.filter(c => c.visible && c.nextReview <= Date.now());
    document.getElementById('total-info').innerText = `Total: ${deck.length} | Pendientes hoy: ${dueCards.length}`;
}

// 2. CAMBIO DE PÁGINA
function startStudy() {
    const dueCards = deck.filter(c => c.visible && c.nextReview <= Date.now());
    if (dueCards.length === 0) return alert("¡No hay cartas pendientes para hoy!");
    
    document.getElementById('setup-view').style.display = 'none';
    document.getElementById('study-view').style.display = 'block';
    showNextCard();
}

// 3. LÓGICA DE ESTUDIO
function showNextCard() {
    const dueCards = deck.filter(c => c.visible && c.nextReview <= Date.now());
    document.getElementById('cards-left').innerText = dueCards.length;

    if (dueCards.length === 0) {
        alert("¡Has terminado por ahora!");
        location.reload();
        return;
    }

    // Buscamos el índice real en el mazo original
    currentCardIndex = deck.indexOf(dueCards[0]);
    showingAnswer = false;
    document.getElementById('card').innerText = deck[currentCardIndex].q;
    document.getElementById('answer-controls').style.display = 'none';
}

document.getElementById('card').addEventListener('click', () => {
    if (currentCardIndex === -1 || showingAnswer) return;
    showingAnswer = true;
    document.getElementById('card').innerText = deck[currentCardIndex].a;
    document.getElementById('answer-controls').style.display = 'block';
});

// 4. CONFIGURAR REPETICIÓN (DÍAS)
function setSchedule(days) {
    const msInDay = 24 * 60 * 60 * 1000;
    
    if (days === 0) {
        // Se queda en la cola de hoy (le ponemos un timestamp muy viejo para que siga saliendo)
        deck[currentCardIndex].nextReview = Date.now();
    } else {
        // Se programa para el futuro
        deck[currentCardIndex].nextReview = Date.now() + (days * msInDay);
    }

    localStorage.setItem('myFlashcards', JSON.stringify(deck));
    showNextCard();
}

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


