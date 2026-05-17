let allDecks = JSON.parse(localStorage.getItem('myFlashcardDecks')) || {};
let allBooks = JSON.parse(localStorage.getItem('myFlashcardBooks')) || {};
let currentDeckName = null; 
let currentCardIndex = -1;
let showingAnswer = false; 
let tempCardsArray = []; // Guarda temporalmente las cartas del Excel antes de confirmar
let isReverseMode = JSON.parse(localStorage.getItem('flashcards_reverse')) || false;
let selectedCardIndex = null; // Para saber qué carta estamos editando/borrando
let isEditing = false;
let currentBookName = null;
let rendezvousBook = null; // Almacenará la instancia activa de epubjs
let bookRendition = null;  // Almacenará el controlador visual de la lectura
let selectedWordFromText = "";
let calculatedAiBackText = "";
let readingMode = 'scroll'; // 'scroll' o 'pages'
let isFullscreenReader = false;
let touchStartX = 0;
let touchEndX = 0;
let currentPageIdx = 0;
let totalPagesCount = 1;
let iaPensando = false;

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
    renderBooks();
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
        container.innerHTML = '<p style="color: #666;">No hay mazos añadidos. Importa un Excel arriba.</p>';
        return;
    }

    deckNames.forEach(name => {
        const deck = allDecks[name];
        const dueCardsCount = deck.filter(c => c.nextReview <= now).length;
        
        const div = document.createElement('div');
        div.className = 'deck-card';
        
        // ACCIÓN PRINCIPAL: Clicar en el mazo te lleva DIRECTAMENTE a ver las cartas
        div.onclick = (e) => {
            // Evitamos que se dispare si el usuario pulsó específicamente alguno de los botones derechos
            if (e.target.closest('button')) return;
            viewDeckList(name);
        };

        div.innerHTML = `
            <div>
                <strong style="font-size: 1.1rem;">${name}</strong><br>
                <span style="font-size: 0.85rem; color: #666;">
                    Total: ${deck.length} | Hoy: <b style="color: ${dueCardsCount > 0 ? '#ff4444' : 'green'};">${dueCardsCount}</b>
                </span>
            </div>
            <div class="deck-actions">
                <!-- Botón directo para estudiar -->
                <button onclick="startStudy('${name}')" style="background: #007bff; padding: 10px 14px; font-weight: bold; margin: 0; font-size: 0.9rem;" ${dueCardsCount === 0 ? 'disabled style="background:#ccc; cursor:not-allowed;"' : ''}>Estudiar</button>
                <!-- Cubo de basura directo para eliminar -->
                <button class="btn-delete-direct" onclick="deleteDeck('${name}')">🗑️</button>
            </div>
        `;
        container.appendChild(div);
    });
}

// RENDERIZADO DE LIBROS ACTUALIZADO
function renderBooks() {
    const container = document.getElementById('books-container');
    if (!container) return;
    container.innerHTML = '';

    const bookNames = Object.keys(allBooks);
    if (bookNames.length === 0) {
        container.innerHTML = '<p style="color: #666;">No hay libros importados. Sube un EPUB arriba.</p>';
        return;
    }

    bookNames.forEach(name => {
        const book = allBooks[name];
        const div = document.createElement('div');
        div.className = 'deck-card';
        div.style.borderColor = '#6f42c1'; // Mantenemos tu distintivo borde lila para lecturas
        
        // ACCIÓN PRINCIPAL: Clicar en el libro abre el lector interactivo
        div.onclick = (e) => {
            if (e.target.closest('button')) return;
            startReadingBook(name);
        };

        div.innerHTML = `
            <div>
                <strong style="font-size: 1.1rem;">📖 ${name}</strong><br>
                <span style="font-size: 0.85rem; color: #666;">
                    Secciones: ${book.chapters ? book.chapters.length : 0}
                </span>
            </div>
            <div class="deck-actions">
                <!-- Cubo de basura directo para eliminar el texto -->
                <button class="btn-delete-direct" onclick="deleteBook('${name}')">🗑️</button>
            </div>
        `;
        container.appendChild(div);
    });
}


// Cambiar de pestaña entre añadir Mazo o añadir Libro
function switchAddTab(type) {
    const formDeck = document.getElementById('form-add-deck');
    const formBook = document.getElementById('form-add-book');
    const tabDeck = document.getElementById('tab-add-deck');
    const tabBook = document.getElementById('tab-add-book');

    if (type === 'deck') {
        formDeck.style.display = 'block';
        formBook.style.display = 'none';
        tabDeck.style.background = '#28a745';
        tabBook.style.background = '#6c757d';
    } else {
        formDeck.style.display = 'none';
        formBook.style.display = 'block';
        tabDeck.style.background = '#6c757d';
        tabBook.style.background = '#28a745';
    }
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

// Función para procesar e importar el archivo EPUB
function importBook() {
    const nameInput = document.getElementById('book-name');
    const fileInput = document.getElementById('book-input');
    const name = nameInput.value.trim();
    const file = fileInput.files[0];

    if (!name || !file) return alert("Selecciona un archivo .epub y asígnale un nombre.");
    if (allBooks[name]) return alert("Ya existe un libro con ese nombre.");

    if (!window.ePub) {
        alert("Cargando el desempaquetador del libro. Por favor, vuelve a pulsar el botón en 3 segundos.");
        injectImportLibraries();
        return;
    }

    const reader = new FileReader();
    reader.onload = async function(e) {
        try {
            const eDoc = ePub(e.target.result);
            await eDoc.opened;
            
            // Intentamos leer el índice visible por si acaso
            const navigation = await eDoc.loaded.navigation;
            let tocMap = {};
            if (navigation && navigation.toc) {
                navigation.toc.forEach(item => {
                    // Limpiamos las rutas para poder cruzarlas con el spine
                    let cleanHref = item.href.split('#')[0].replace('../', '');
                    tocMap[cleanHref] = item.label ? item.label.trim() : null;
                });
            }

            let chaptersData = [];
            let chapterCounter = 1;

            // SOLUCIÓN: Recorremos el Spine (Columna vertebral obligatoria) en vez del TOC
            // Esto garantiza que procesamos el 100% de las páginas del libro
            for (let i = 0; i < eDoc.spine.items.length; i++) {
                const section = eDoc.spine.items[i];
                if (section) {
                    await section.load(eDoc.load.bind(eDoc));
                    
                    let htmlContent = section.document.body.innerHTML;
                    
                    // Limpieza opcional: Eliminamos imágenes pesadas para no saturar el LocalStorage
                    htmlContent = htmlContent.replace(/<img[^>]*>/g, '🖼️ [Imagen]');

                    // Extraemos el texto plano para comprobar si la sección está vacía
                    let tempDiv = document.createElement('div');
                    tempDiv.innerHTML = htmlContent;
                    let plainText = tempDiv.innerText.trim();

                    // Solo guardamos la sección si realmente contiene texto para leer
                    if (plainText.length > 30) {
                        // Intentamos ponerle el nombre real del capítulo usando nuestro mapa del TOC
                        let cleanSectionHref = section.href.replace('../', '');
                        let title = tocMap[cleanSectionHref] || `Sección ${chapterCounter++}`;

                        chaptersData.push({
                            title: title,
                            html: htmlContent
                        });
                    }
                    section.unload();
                }
            }

            if (chaptersData.length === 0) {
                throw new Error("El libro no contiene secciones de texto compatibles.");
            }

            allBooks[name] = {
                title: name,
                chapters: chaptersData,
                lastChapterIndex: 0
            };

            localStorage.setItem('myFlashcardBooks', JSON.stringify(allBooks));
            nameInput.value = '';
            fileInput.value = '';
            toggleMenu('add-deck-menu');
            renderBooks();
            alert(`¡"${name}" importado con éxito! Se han extraído ${chaptersData.length} secciones de lectura.`);

        } catch (err) {
            console.error("Fallo en descompresión:", err);
            alert("Este EPUB tiene un formato protegido (DRM) o incompatible. Intenta convertirlo a 'EPUB fluido' con Calibre antes de subirlo.");
        }
    };
    reader.readAsArrayBuffer(file);
}


// Inyección segura en caliente
function injectImportLibraries() {
    const s1 = document.createElement('script');
    s1.src = "https://cloudflare.com";
    const s2 = document.createElement('script');
    s2.src = "https://cloudflare.com";
    document.head.appendChild(s1);
    setTimeout(() => document.head.appendChild(s2), 500);
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

// Inicia la pantalla del lector cargando el libro de memoria
function startReadingBook(name) {
    currentBookName = name;
    const bookData = allBooks[name];

    document.getElementById('setup-view').style.display = 'none';
    document.getElementById('book-view').style.display = 'block';
    history.pushState({view: 'book-reader'}, "");

    if (!bookData.lastChapterIndex) bookData.lastChapterIndex = 0;

    // Rellenamos e inicializamos el selector visual del índice superior
    updateChapterSelectUI();

    // Forzar modo scroll por defecto al iniciar y restaurar pantalla normal
    readingMode = 'scroll';
    isFullscreenReader = false;
    applyInterfaceLayout();
    renderCurrentChapterText();
    
    // Configurar el detector de gestos táctiles (Swipe) sobre el contenedor
    setupSwipeGestures();

    // --- CORRECCIÓN: EVENTO DE SCROLL AUTOMÁTICO ENTRE CAPÍTULOS ---
    const viewer = document.getElementById('book-viewer-container');
    viewer.onscroll = function() {
        if (readingMode !== 'scroll') return; // Solo actúa en modo scroll vertical

        const bData = allBooks[currentBookName];
        const currentIdx = bData.lastChapterIndex;

        // Si el usuario llega al fondo del capítulo (margen de 5px de tolerancia)
        if (viewer.scrollTop + viewer.clientHeight >= viewer.scrollHeight - 5) {
            if (currentIdx < bData.chapters.length - 1) {
                changeChapter(currentIdx + 1); // Pasa al siguiente capítulo automáticamente
            }
        }
        // Si el usuario sube arriba del todo e intenta seguir arrastrando hacia arriba
        else if (viewer.scrollTop <= 0) {
            if (currentIdx > 0) {
                changeChapter(currentIdx - 1); // Vuelve al capítulo anterior automáticamente
                // Forzamos el scroll abajo del todo para que la lectura continúe fluida
                setTimeout(() => { viewer.scrollTop = viewer.scrollHeight; }, 50);
            }
        }
    };

    // Asignamos los controles inferiores a nuestra lógica de navegación inteligente
    document.getElementById('btn-prev-page').onclick = () => handlePageNavigation('prev');
    document.getElementById('btn-next-page').onclick = () => handlePageNavigation('next');
}

// Configura la visualización limpia del HTML según el modo elegido
function renderCurrentChapterText() {
    const bookData = allBooks[currentBookName];
    const chapter = bookData.chapters[bookData.lastChapterIndex];
    const bookArea = document.getElementById('book-area');
    const viewer = document.getElementById('book-viewer-container');
    
    bookArea.innerHTML = chapter.html;
    currentPageIdx = 0;

    if (readingMode === 'pages') {
        // Truco CSS Avanzado: Convertimos el div en un periódico que crece hacia la derecha
        viewer.style.overflowY = 'hidden';
        viewer.style.overflowX = 'hidden';
        bookArea.style.cssText = `
            column-width: ${viewer.clientWidth - 10}px;
            column-gap: 20px;
            height: ${viewer.clientHeight - 10}px;
            transition: transform 0.3s cubic-bezier(0.25, 1, 0.5, 1);
            transform: translateX(0px);
        `;
        
        // Calculamos cuántas páginas se han generado dinámicamente según el volumen de texto
        setTimeout(() => {
            const totalWidth = bookArea.scrollWidth;
            const colWidth = viewer.clientWidth + 10;
            totalPagesCount = Math.max(1, Math.ceil(totalWidth / colWidth));
            updatePageCounter();
        }, 100);
    } else {
        // Restauramos el scroll vertical normal
        viewer.style.overflowY = 'auto';
        viewer.style.overflowX = 'hidden';
        bookArea.style.cssText = "width: 100%; height: auto;";
        document.getElementById('page-counter-label').innerText = "Scroll";
    }

    bookArea.onclick = function(e) {
        if (isFullscreenReader) {
            triggerBarsIndicator();
        }
    };

    // Activador por doble click para la IA de vocabulario
    bookArea.ondblclick = function(e) {
        if (iaPensando) {
            alert("Espera a que la IA termine con la palabra anterior.");
            return;
        }
        let range, textNode, offset;
        if (document.caretPositionFromPoint) {
            let pos = document.caretPositionFromPoint(e.clientX, e.clientY);
            if (pos) { textNode = pos.offsetNode; offset = pos.offset; }
        } else if (document.caretRangeFromPoint) {
            range = document.caretRangeFromPoint(e.clientX, e.clientY);
            if (range) { textNode = range.startContainer; offset = range.startOffset; }
        }

        if (!textNode || textNode.nodeType !== Node.TEXT_NODE) return;
        const data = textNode.data;
        
        let start = offset;
        while (start > 0 && !/\s/.test(data[start - 1])) start--;
        let end = offset;
        while (end < data.length && !/\s/.test(data[end])) end++;

        let selectedText = data.substring(start, end).trim();
        if (!selectedText || selectedText.includes(" ") || selectedText.length < 2) return;

        let cleanWord = selectedText.replace(/[.,\/#!$%\^&\*;:{}=\-_`~()?"'¿¡]/g,"");
        let contextText = textNode.parentElement ? (textNode.parentElement.innerText || "") : "";
        selectedWordFromText = cleanWord;

        iaPensando = true;
        openReaderPopup(cleanWord, contextText);
    };
}

// Alternar entre modo scroll vertical u horizontal por páginas
function toggleReadingMode() {
    readingMode = (readingMode === 'scroll') ? 'pages' : 'scroll';
    document.getElementById('btn-toggle-mode').innerText = (readingMode === 'scroll') ? "📜 Scroll" : "📖 Páginas";
    applyInterfaceLayout();
    renderCurrentChapterText();
}

// Controlar la navegación física de las columnas en modo Páginas
function navigatePage(direction) {
    if (readingMode !== 'pages') return;
    const viewer = document.getElementById('book-viewer-container');
    const bookArea = document.getElementById('book-area');
    const step = viewer.clientWidth + 10;

    if (direction === 'next' && currentPageIdx < totalPagesCount - 1) {
        currentPageIdx++;
    } else if (direction === 'prev' && currentPageIdx > 0) {
        currentPageIdx--;
    }

    bookArea.style.transform = `translateX(-${currentPageIdx * step}px)`;
    updatePageCounter();
}

function updatePageCounter() {
    document.getElementById('page-counter-label').innerText = `Pág ${currentPageIdx + 1} / ${totalPagesCount}`;
}

// LÓGICA DE PANTALLA COMPLETA (MODO INMERSIVO)
function toggleFullscreenReader() {
    isFullscreenReader = !isFullscreenReader;
    const btn = document.getElementById('btn-fullscreen-toggle');
    
    if (isFullscreenReader) {
        btn.innerText = "🛜"; 
        btn.style.background = "#dc3545"; 
    } else {
        btn.innerText = "⛶"; 
        btn.style.background = "#28a745"; 
    }
    
    applyInterfaceLayout();
}

function triggerBarsIndicator() {
    const topBar = document.getElementById('book-top-bar');
    topBar.style.display = 'flex';
    
    clearTimeout(window.barsTimeout);
    window.barsTimeout = setTimeout(() => {
        if (isFullscreenReader) {
            topBar.style.display = 'none';
            applyInterfaceLayout();
        }
    }, 3000);
}

function applyInterfaceLayout() {
    const topBar = document.getElementById('book-top-bar');
    const bottomBar = document.getElementById('book-bottom-pagination');
    const viewer = document.getElementById('book-viewer-container');

    if (isFullscreenReader) {
        topBar.style.display = 'none';
        bottomBar.style.display = (readingMode === 'pages') ? 'flex' : 'none';
        viewer.style.height = (readingMode === 'pages') ? "calc(100vh - 65px)" : "100vh";
    } else {
        topBar.style.display = 'flex';
        bottomBar.style.display = (readingMode === 'pages') ? 'flex' : 'none';
        viewer.style.height = (readingMode === 'pages') ? "calc(100vh - 120px)" : "calc(100vh - 75px)";
    }
}

// CAPTURA DE GESTOS SWIPE (Deslizar el dedo para pasar páginas)
function setupSwipeGestures() {
    const viewer = document.getElementById('book-viewer-container');
    if (!viewer) return;
    
    viewer.addEventListener('touchstart', e => {
        touchStartX = e.changedTouches[0].screenX;
    }, {passive: true});

    viewer.addEventListener('touchend', e => {
        touchEndX = e.changedTouches[0].screenX;
        handleSwipeLogic();
    }, {passive: true});
}

function handleSwipeLogic() {
    if (readingMode !== 'pages') return;
    const swipeDistance = touchStartX - touchEndX;
    
    if (swipeDistance > 50) {
        handlePageNavigation('next'); 
    } else if (swipeDistance < -50) {
        handlePageNavigation('prev'); 
    }
}

// Lógica unificada para el paso de hojas y capítulos
function handlePageNavigation(direction) {
    const viewer = document.getElementById('book-viewer-container');
    const bookData = allBooks[currentBookName];
    const currentIdx = bookData.lastChapterIndex;

    if (readingMode === 'pages') {
        if (direction === 'next') {
            if (currentPageIdx < totalPagesCount - 1) {
                navigatePage('next');
            } else if (currentIdx < bookData.chapters.length - 1) {
                changeChapter(currentIdx + 1); // Capítulo siguiente automático
            } else {
                alert("Has llegado al final del libro. 🎉");
            }
        } else {
            if (currentPageIdx > 0) {
                navigatePage('prev');
            } else if (currentIdx > 0) {
                // Retrocedemos de capítulo y nos posicionamos al final de sus páginas
                changeChapter(currentIdx - 1);
                setTimeout(() => {
                    currentPageIdx = totalPagesCount - 1;
                    const step = viewer.clientWidth + 10;
                    document.getElementById('book-area').style.transform = `translateX(-${currentPageIdx * step}px)`;
                    updatePageCounter();
                }, 150);
            }
        }
    } else {
        // MODALIDAD ADAPTADA PARA LECTURA VERTICAL (Scroll botones inferiores)
        if (direction === 'next') {
            if (viewer.scrollTop + viewer.clientHeight >= viewer.scrollHeight - 20) {
                if (currentIdx < bookData.chapters.length - 1) {
                    changeChapter(currentIdx + 1);
                } else {
                    alert("Has llegado al final del libro. 🎉");
                }
            } else {
                viewer.scrollTop += 300;
            }
        } else {
            if (viewer.scrollTop <= 5) {
                if (currentIdx > 0) {
                    changeChapter(currentIdx - 1);
                    setTimeout(() => viewer.scrollTop = viewer.scrollHeight, 100);
                }
            } else {
                viewer.scrollTop -= 300;
            }
        }
    }
}

// --- CORRECCIÓN: CAMBIO DE CAPÍTULO Y ACTUALIZACIÓN DEL ÍNDICE SUPERIOR ---
function changeChapter(index) {
    const idx = parseInt(index);
    allBooks[currentBookName].lastChapterIndex = idx;
    localStorage.setItem('myFlashcardBooks', JSON.stringify(allBooks));
    
    // Forzamos al selector a reflejar visualmente la sección correcta
    updateChapterSelectUI();
    
    renderCurrentChapterText();
    document.getElementById('book-viewer-container').scrollTop = 0; 
}

// Función auxiliar encargada de redibujar y marcar el capítulo activo en el select
function updateChapterSelectUI() {
    const bookData = allBooks[currentBookName];
    const select = document.getElementById('book-chapter-select');
    if (!select || !bookData) return;

    select.innerHTML = '';
    bookData.chapters.forEach((ch, idx) => {
        const opt = document.createElement('option');
        opt.value = idx;
        opt.innerText = ch.title;
        if (idx === bookData.lastChapterIndex) opt.selected = true; // Sincroniza la UI
        select.appendChild(opt);
    });
}


async function openReaderPopup(word, context) {
    const apiKey = localStorage.getItem('gemini_api_key');
    const popup = document.getElementById('reader-popup');
    const loading = document.getElementById('popup-loading');
    const resultArea = document.getElementById('popup-result-area');
    const select = document.getElementById('popup-deck-select');

    document.getElementById('popup-word').innerText = word;
    popup.style.display = 'flex';
    loading.style.display = 'block';
    resultArea.style.display = 'none';

    // Rellenamos el selector con los mazos actuales del LocalStorage
    select.innerHTML = '';
    const deckNames = Object.keys(allDecks);
    if (deckNames.length === 0) {
        const opt = document.createElement('option');
        opt.innerText = "No tienes mazos. Crea uno primero.";
        select.appendChild(opt);
    } else {
        deckNames.forEach(name => {
            const opt = document.createElement('option');
            opt.value = name; opt.innerText = name;
            select.appendChild(opt);
        });
    }

    if (!apiKey) {
        document.getElementById('popup-definition').value = "Configura tu clave de API en Ajustes para traducir.";
        loading.style.display = 'none';
        resultArea.style.display = 'block';
        return;
    }

    // PROMPT AVANZADO CON CONTEXTO COMPLETO
    const prompt = `Analiza la palabra "${word}" basándote exactamente en el contexto de esta frase/párrafo: "${context}".
Detecta el idioma de la lectura, pero no lo escribas. Proporciona una definición corta y precisa de la palabra en ese mismo idioma, incluyendo ademas (separados con ';') otras acepciones que tenga, si son importantes.
Luego, entre paréntesis, incluye un par de sinónimos usando el formato (=sinónimo1, sinónimo2).
Finalmente, añade un guion y su traducción exacta al español.
Devuelve ÚNICAMENTE el resultado final en una sola línea, imitando estrictamente este formato:
a flat surface for storage (=ledge, rack) - Estante`;

    const apiEndpoint = "https://generativelanguage.googleapis.com/v1/models/gemini-2.5-flash:generateContent?key=" + apiKey;
    const payload = { contents: [{ parts: [{ text: prompt }] }] };

    try {
        const response = await fetch(apiEndpoint, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(payload)
        });
        const data = await response.json();

        if (data.candidates && data.candidates[0] && data.candidates[0].content && data.candidates[0].content.parts && data.candidates[0].content.parts[0]) {
            calculatedAiBackText = data.candidates[0].content.parts[0].text.trim().replace(/\n/g, '');
            document.getElementById('popup-definition').value = calculatedAiBackText;
        } else {
            throw new Error();
        }
    } catch (e) {
        document.getElementById('popup-definition').value = "Error al conectar con Gemini.";
    } finally {
        loading.style.display = 'none';
        resultArea.style.display = 'block';
        iaPensando = false;
    }
}

function addPopupCardToDeck() {
    const deckName = document.getElementById('popup-deck-select').value;
    const finalDefinition = document.getElementById('popup-definition').value.trim();

    if (!deckName || !allDecks[deckName]) return alert("Selecciona un mazo válido.");
    if (!finalDefinition) return alert("Espera a recibir la definición de la IA.");

    // Inyectamos la carta directamente en el mapa global del mazo
    allDecks[deckName].push({
        q: selectedWordFromText,
        a: finalDefinition,
        nextReview: 0 // Lista para estudiar hoy mismo
    });

    localStorage.setItem('myFlashcardDecks', JSON.stringify(allDecks));
    alert(`¡Carta "${selectedWordFromText}" guardada con éxito en el mazo "${deckName}"!`);
    closeReaderPopup();
}

function closeReaderPopup() {
    document.getElementById('reader-popup').style.display = 'none';
}

// Salir del lector y regresar a la pantalla principal
function goBackFromBook() {
    document.getElementById('book-view').style.display = 'none';
    document.getElementById('setup-view').style.display = 'block';
    renderBooks();
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

// Eliminar un libro
function deleteBook(name) {
    if (confirm(`¿Seguro que quieres eliminar el libro "${name}"?`)) {
        delete allBooks[name];
        localStorage.setItem('myFlashcardBooks', JSON.stringify(allBooks));
        renderBooks();
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
    const apiEndpoint = "https://generativelanguage.googleapis.com/v1/models/gemini-2.5-flash:generateContent?key=" + apiKey;

    const payload = {
        contents: [{ parts: [{ text: prompt }] }]
    };
    
    try {
        /* SOLUCIÓN DEFINITIVA: Convertimos el JSON en un Blob plano de texto.
        const response = await fetch(apiEndpoint, {
            method: 'POST',
            body: new Blob([JSON.stringify(payload)], { type: 'text/plain' })
        });
        */
        const response = await fetch(apiEndpoint, {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json'
            },
            body: JSON.stringify(payload)
        });
        if (!response.ok) {
            const errorData = await response.json();
            console.error("Error de la API:", errorData);
            throw new Error(errorData.error?.message || "Error en la petición");
        }
        
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
    const popupElement = document.getElementById('reader-popup');

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
    if (document.getElementById('book-view').style.display === 'block') {
        if (isFullscreenReader) {
            exitFullscreenReader();
        } else {
            goBackFromBook();
        }
        return;
    }
    if(popupElement && popupElement.style.display === 'flex') {
        closeReaderPopup();
        return;
    }
    // Solo si no había ningún menú abierto, volvemos a la pantalla principal
    goToHome();
};


