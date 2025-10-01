'use strict';

// URL вашего веб-приложения Google Apps Script
const SCRIPT_URL = "https://script.google.com/macros/s/AKfycbyTkziw2S7sA6i-3FX0bOZjpi2cLT1iSoN9-3BgWV0JdeFi1RSMyJQbdpWAH8BMD_OWpg/exec"; // <-- НЕ ЗАБУДЬТЕ ВСТАВИТЬ СВОЙ URL

// Получаем ссылки на наши "экраны"
const gridView = document.getElementById('quiz-grid-view');
const questionView = document.getElementById('question-view');
const loadingIndicator = document.getElementById('loading');

Office.onReady(info => {
    if (info.host === Office.HostType.PowerPoint) {
        loadingIndicator.textContent = "Загрузка...";
        fetchQuizData();
        
        // Назначаем обработчик на кнопку "Назад к таблице"
        document.getElementById('back-to-grid').onclick = showGridView;
    }
});

async function fetchQuizData() {
    try {
        const response = await fetch(SCRIPT_URL);
        if (!response.ok) throw new Error(`HTTP error! status: ${response.status}`);
        
        const result = await response.json();
        if (result.status === "success") {
            buildQuizGrid(result.data);
        } else {
            throw new Error(result.message || "Ошибка при получении данных.");
        }
    } catch (error) {
        console.error("Ошибка при загрузке данных квиза:", error);
        loadingIndicator.textContent = `Ошибка: ${error.message}`;
    }
}

function buildQuizGrid(data) {
    const container = document.getElementById("quiz-container");
    container.innerHTML = "";
    loadingIndicator.style.display = "none";

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
                
                // Сохраняем все данные вопроса в data-атрибутах
                cell.dataset.category = category;
                cell.dataset.question = q.question;
                cell.dataset.answers = JSON.stringify(q.answers);
                cell.dataset.correctAnswerIndex = q.correctAnswerIndex;
                
                cell.onclick = () => handleQuestionClick(cell);
                container.appendChild(cell);
            });
        }
    }
}

function handleQuestionClick(cell) {
    // Получаем данные из ячейки
    const category = cell.dataset.category;
    const questionText = cell.dataset.question;
    const answers = JSON.parse(cell.dataset.answers);
    
    // Заполняем "экран" вопроса
    document.getElementById('question-category').textContent = `${category} за ${cell.textContent}`;
    document.getElementById('question-text').textContent = questionText;
    
    const answersContainer = document.getElementById('answers-container');
    answersContainer.innerHTML = ''; // Очищаем предыдущие ответы
    
    answers.forEach((answer, index) => {
        const button = document.createElement('button');
        button.className = 'answer-btn';
        button.textContent = `${String.fromCharCode(65 + index)}) ${answer}`;
        button.onclick = () => {
            // Здесь будет логика проверки ответа
            console.log(`Выбран ответ: ${answer}`);
            // После ответа можно, например, вернуться к сетке
            showGridView();
        };
        answersContainer.appendChild(button);
    });

    // Отключаем ячейку, чтобы ее нельзя было выбрать снова
    cell.classList.add("disabled");
    
    // Переключаем экраны
    showQuestionView();
}

// Функции для переключения видимости "экранов"
function showGridView() {
    questionView.style.display = 'none';
    gridView.style.display = 'block';
}

function showQuestionView() {
    gridView.style.display = 'none';
    questionView.style.display = 'block';
}
