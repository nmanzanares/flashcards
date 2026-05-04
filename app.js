// Registro del Service Worker (necesario para PWA)
if ('serviceWorker' in navigator) {
    navigator.serviceWorker.register('sw.js');
}

// Datos de ejemplo (luego podrías hacer un formulario para añadirlas)
const deck = [
    { q: "¿Capital de Francia?", a: "París" },
    { q: "¿2 + 2?", a: "4" },
    { q: "¿Color del cielo?", a: "Azul" }
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
    reader.onload = function(e) {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        
        // Obtener la primera hoja
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];
        
        // Convertir a JSON (asumiendo que la fila 1 son encabezados como "pregunta" y "respuesta")
        const jsonData = XLSX.utils.sheet_to_json(worksheet);
        
        // Mapear los datos a tu formato de flashcards
        const newCards = jsonData.map(row => ({
            q: row.pregunta || row.Pregunta, // Busca la columna por nombre
            a: row.respuesta || row.Respuesta
        }));

        // Guardar en el mazo actual y en localStorage
        deck.push(...newCards);
        localStorage.setItem('myFlashcards', JSON.stringify(deck));
        
        alert(`¡Se han importado ${newCards.length} tarjetas con éxito!`);
    };
    reader.readAsArrayBuffer(file);
});
