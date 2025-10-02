'use strict';

const SCRIPT_URL = "https://script.google.com/macros/s/AKfycbyTkziw2S7sA6i-3FX0bOZjpi2cLT1iSoN9-3BgWV0JdeFi1RSMyJQbdpWAH8BMD_OWpg/exec"; // <-- НЕ ЗАБУДЬТЕ ВСТАВИТЬ СВОЙ URL

const gridView = document.getElementById('quiz-grid-view');
const questionView = document.getElementById('question-view');
const loadingIndicator = document.getElementById('loading');

let currentQuestionCell = null;

Office.onReady(info => {
    if (info.host === Office.HostType.PowerPoint) {
        loadingIndicator.textContent = "Загрузка...";
        fetchQuizData();
        document.getElementById('back-to-grid').onclick = showGridView;
    }
});

async function fetchQuizData() {
    try {
        // ИЗМЕНЕНИЕ: Добавляем параметр action=getQuizData к URL
        const response = await fetch(`${SCRIPT_URL}?action=getQuizData`);
        if (!response.ok) throw new Error(`HTTP error! status: ${response.status}`);
        const result = await response.json();
        if (result.status === "success") {
            buildQuizTable(result.data);
        } else {
            throw new Error(result.message || "Ошибка при получении данных.");
        }
    } catch (error) {
        console.error("Ошибка:", error);
        loadingIndicator.textContent = `Ошибка: ${error.message}`;
    }
}

// Функция buildQuizTable остается без изменений
function buildQuizTable(data) {
    const table = document.getElementById("quiz-table");
    table.innerHTML = "";
    loadingIndicator.style.display = "none";
    const pointValues = [...new Set(Object.values(data).flat().map(q => q.points))].sort((a, b) => a - b);
    const thead = table.createTHead();
    const headerRow = thead.insertRow();
    headerRow.insertCell().textContent = "Категория";
    pointValues.forEach(points => {
        const th = document.createElement('th');
        th.textContent = points;
        headerRow.appendChild(th);
    });
    const tbody = table.createTBody();
    for (const category in data) {
        if (data.hasOwnProperty(category)) {
            const row = tbody.insertRow();
            row.insertCell().textContent = category;
            row.cells[0].className = 'category-name-cell';
            let isFirstQuestionInRow = true;
            const questionsInCategory = data[category];
            pointValues.forEach(points => {
                const cell = row.insertCell();
                const question = questionsInCategory.find(q => q.points === points);
                if (question) {
                    cell.textContent = question.points;
                    cell.className = 'question-cell';
                    cell.dataset.category = category;
                    cell.dataset.question = question.question;
                    cell.dataset.answers = JSON.stringify(question.answers);
                    cell.dataset.correctAnswerIndex = question.correctAnswerIndex;
                    cell.dataset.hint = question.hint;
                    cell.dataset.points = question.points;
                    cell.onclick = () => handleQuestionClick(cell);
                    if (isFirstQuestionInRow) {
                        isFirstQuestionInRow = false;
                    } else {
                        cell.classList.add('locked');
                    }
                } else {
                    cell.className = 'placeholder-cell';
                }
            });
        }
    }
}

function handleQuestionClick(cell) {
    currentQuestionCell = cell;
    const { category, question, answers: answersJson, hint, points } = cell.dataset;
    const answers = JSON.parse(answersJson);

    // Активируем вопрос для игроков
    const questionDataForPlayers = {
        category,
        question,
        points,
        answers: answers.map((ans, index) => String.fromCharCode(65 + index)) // Отправляем только буквы A, B, C...
    };
    postToServer({ action: 'setActiveQuestion', data: questionDataForPlayers });

    // Отображаем вопрос для ведущего
    document.getElementById('question-category').textContent = `${category} за ${points}`;
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
    const hintContainer = document.getElementById('hint-container');
    const showHintBtn = document.getElementById('show-hint-btn');
    if (hint) {
        document.getElementById('hint-text').textContent = hint;
        showHintBtn.style.display = 'block';
        showHintBtn.onclick = () => {
            hintContainer.style.display = 'block';
            showHintBtn.style.display = 'none';
        };
    } else {
        showHintBtn.style.display = 'none';
    }
    hintContainer.style.display = 'none';
    showQuestionView();
}

function checkAnswer(selectedIndex) {
    const correctIndex = parseInt(currentQuestionCell.dataset.correctAnswerIndex, 10);
    const answerButtons = document.querySelectorAll('#answers-container .answer-btn');
    answerButtons.forEach(btn => btn.disabled = true);
    if (selectedIndex === correctIndex) {
        answerButtons[selectedIndex].classList.add('correct');
    } else {
        answerButtons[selectedIndex].classList.add('incorrect');
        if (correctIndex > -1) answerButtons[correctIndex].classList.add('correct');
    }
    currentQuestionCell.classList.add("disabled");
    let nextCell = currentQuestionCell.nextElementSibling;
    while(nextCell && !nextCell.classList.contains('question-cell')) {
        nextCell = nextCell.nextElementSibling;
    }
    if (nextCell) {
        nextCell.classList.remove('locked');
    }
    // Убрали авто-возврат, теперь ведущий нажимает кнопку "Назад" сам
}

function showGridView() {
    // Сбрасываем активный вопрос для игроков
    postToServer({ action: 'setActiveQuestion', data: null });
    questionView.style.display = 'none';
    gridView.style.display = 'block';
}

function showQuestionView() {
    gridView.style.display = 'none';
    questionView.style.display = 'block';
}

// Новая универсальная функция для отправки POST-запросов на сервер
async function postToServer(payload) {
    try {
        await fetch(SCRIPT_URL, {
            method: 'POST',
            mode: 'no-cors', // Важно для отправки данных в Google Apps Script из надстройки
            cache: 'no-cache',
            headers: {
                'Content-Type': 'application/json'
            },
            body: JSON.stringify(payload)
        });
    } catch (error) {
        console.error('Ошибка при отправке данных на сервер:', error);
    }
}
