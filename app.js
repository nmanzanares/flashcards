// Registro del Service Worker (necesario para PWA)
if ('serviceWorker' in navigator) {
    navigator.serviceWorker.register('sw.js');
}

// Datos de ejemplo (luego podrías hacer un formulario para añadirlas)
const deck = [
    { q: "¿Capital de Francia?", a: "París" },
    { q: "¿2 + 2?", a: "5" },
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


