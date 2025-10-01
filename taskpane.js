'use strict';

// URL вашего веб-приложения Google Apps Script
const SCRIPT_URL = "https://script.google.com/macros/s/AKfycbyTkziw2S7sA6i-3FX0bOZjpi2cLT1iSoN9-3BgWV0JdeFi1RSMyJQbdpWAH8BMD_OWpg/exec"; // <-- НЕ ЗАБУДЬТЕ ВСТАВИТЬ СВОЙ URL

// Получаем ссылки на DOM-элементы
const gridView = document.getElementById('quiz-grid-view');
const questionView = document.getElementById('question-view');
const loadingIndicator = document.getElementById('loading');
const scoreDisplay = document.getElementById('score-display');

let score = 0;
let currentQuestionCell = null; // Сохраняем ячейку текущего вопроса

Office.onReady(info => {
    if (info.host === Office.HostType.PowerPoint) {
        loadingIndicator.textContent = "Загрузка...";
        fetchQuizData();
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
        console.error("Ошибка:", error);
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
                cell.dataset.category = category;
                cell.dataset.question = q.question;
                cell.dataset.answers = JSON.stringify(q.answers);
                cell.dataset.correctAnswerIndex = q.correctAnswerIndex;
                cell.dataset.hint = q.hint;
                cell.dataset.points = q.points;
                cell.onclick = () => handleQuestionClick(cell);
                container.appendChild(cell);
            });
        }
    }
}

function handleQuestionClick(cell) {
    currentQuestionCell = cell; // Сохраняем текущую ячейку
    const { category, question, answers: answersJson, hint } = cell.dataset;
    const answers = JSON.parse(answersJson);

    document.getElementById('question-category').textContent = `${category} за ${cell.textContent}`;
    document.getElementById('question-text').textContent = question;
    
    const answersContainer = document.getElementById('answers-container');
    answersContainer.innerHTML = '';
    
    answers.forEach((answer, index) => {
        const button = document.createElement('button');
        button.className = 'answer-btn';
        button.textContent = `${String.fromCharCode(65 + index)}) ${answer}`;
        button.onclick = () => checkAnswer(index);
        answersContainer.appendChild(button);
    });

    // Логика подсказки
    const hintContainer = document.getElementById('hint-container');
    const hintText = document.getElementById('hint-text');
    const showHintBtn = document.getElementById('show-hint-btn');
    
    if (hint) {
        hintText.textContent = hint;
        showHintBtn.style.display = 'block';
        showHintBtn.onclick = () => {
            hintContainer.style.display = 'block';
            showHintBtn.style.display = 'none'; // Скрываем кнопку после использования
        };
    } else {
        showHintBtn.style.display = 'none';
    }
    hintContainer.style.display = 'none'; // Скрываем подсказку при показе нового вопроса

    showQuestionView();
}

function checkAnswer(selectedIndex) {
    const correctIndex = parseInt(currentQuestionCell.dataset.correctAnswerIndex, 10);
    const points = parseInt(currentQuestionCell.dataset.points, 10);
    const answerButtons = document.querySelectorAll('#answers-container .answer-btn');

    // Блокируем все кнопки
    answerButtons.forEach(btn => btn.disabled = true);

    if (selectedIndex === correctIndex) {
        // Правильный ответ
        answerButtons[selectedIndex].classList.add('correct');
        score += points;
        updateScore();
    } else {
        // Неправильный ответ
        answerButtons[selectedIndex].classList.add('incorrect');
        if (correctIndex > -1) {
            // Подсвечиваем правильный, если он есть
            answerButtons[correctIndex].classList.add('correct');
        }
    }

    // Отключаем ячейку вопроса в сетке
    currentQuestionCell.classList.add("disabled");

    // Возвращаемся к сетке через 2.5 секунды
    setTimeout(showGridView, 2500);
}

function updateScore() {
    scoreDisplay.textContent = `Счет: ${score}`;
}

function showGridView() {
    questionView.style.display = 'none';
    gridView.style.display = 'block';
}

function showQuestionView() {
    gridView.style.display = 'none';
    questionView.style.display = 'block';
}
