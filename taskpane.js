'use strict';

// URL вашего веб-приложения Google Apps Script
const SCRIPT_URL = "https://script.google.com/macros/s/AKfycbyTkziw2S7sA6i-3FX0bOZjpi2cLT1iSoN9-3BgWV0JdeFi1RSMyJQbdpWAH8BMD_OWpg/exec"; // <-- НЕ ЗАБУДЬТЕ ВСТАВИТЬ СВОЙ URL

Office.onReady(info => {
    if (info.host === Office.HostType.PowerPoint) {
        document.getElementById("loading").textContent = "Загрузка...";
        fetchQuizData();
    }
});

async function fetchQuizData() {
    try {
        const response = await fetch(SCRIPT_URL);
        if (!response.ok) {
            throw new Error(`HTTP error! status: ${response.status}`);
        }
        const result = await response.json();

        if (result.status === "success") {
            buildQuizGrid(result.data);
        } else {
            throw new Error(result.message || "Ошибка при получении данных.");
        }
    } catch (error) {
        console.error("Ошибка при загрузке данных квиза:", error);
        document.getElementById("loading").textContent = `Ошибка: ${error.message}`;
    }
}

function buildQuizGrid(data) {
    const container = document.getElementById("quiz-container");
    container.innerHTML = ""; // Очищаем контейнер
    document.getElementById("loading").style.display = "none"; // Скрываем индикатор загрузки

    for (const category in data) {
        if (data.hasOwnProperty(category)) {
            const header = document.createElement("div");
            header.className = "category-header";
            header.textContent = category;
            container.appendChild(header);

            const questions = data[category];
            questions.forEach(q => {
                const cell = document.createElement("div");
                cell.className = "question-cell";
                cell.textContent = q.points;
                cell.dataset.question = q.question;
                cell.dataset.answers = JSON.stringify(q.answers);
                cell.dataset.correctAnswerIndex = q.correctAnswerIndex;
                
                cell.onclick = () => handleQuestionClick(cell, q);
                container.appendChild(cell);
            });
        }
    }
}

function handleQuestionClick(cell, questionData) {
    console.log("Выбран вопрос:", questionData.question);
    
    // Отключаем ячейку после клика, чтобы ее нельзя было выбрать снова
    cell.classList.add("disabled");

    // Форматируем варианты ответов для вывода на слайд (A, B, C, D)
    const formattedAnswers = questionData.answers.map((answer, index) => {
        // Используем буквы латинского алфавита, начиная с 'A'
        const letter = String.fromCharCode(65 + index); 
        return `${letter}) ${answer}`;
    }).join('\n');

    // Формируем итоговый текст для вставки на слайд
    const textToWrite = `Вопрос:\n${questionData.question}\n\nВарианты ответа:\n${formattedAnswers}`;

    // Вставляем текст в текущее выделенное место на слайде (или создаем новое текстовое поле)
    Office.context.document.setSelectedDataAsync(
        textToWrite,
        { coercionType: Office.CoercionType.Text },
        result => {
            if (result.status === Office.AsyncResultStatus.Failed) {
                console.error('Ошибка при вставке текста на слайд:', result.error.message);
            }
        }
    );
}
