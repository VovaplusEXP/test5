'use strict';

const SCRIPT_URL = "https://script.google.com/macros/s/AKfycbyTkziw2S7sA6i-3FX0bOZjpi2cLT1iSoN9-3BgWV0JdeFi1RSMyJQbdpWAH8BMD_OWpg/exec"; // <-- НЕ ЗАБУДЬТЕ ВСТАВИТЬ СВОЙ URL

const gridView = document.getElementById('quiz-grid-view');
const questionView = document.getElementById('question-view');
const loadingIndicator = document.getElementById('loading');
const scoreDisplay = document.getElementById('score-display');

let score = 0;
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
        const response = await fetch(SCRIPT_URL);
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

function buildQuizTable(data) {
    const table = document.getElementById("quiz-table");
    table.innerHTML = "";
    loadingIndicator.style.display = "none";

    // 1. Собрать все уникальные значения баллов для заголовка
    const pointValues = [...new Set(Object.values(data).flat().map(q => q.points))].sort((a, b) => a - b);

    // 2. Создать заголовок таблицы (thead)
    const thead = table.createTHead();
    const headerRow = thead.insertRow();
    headerRow.insertCell().textContent = "Категория"; // Пустая ячейка для названий категорий
    pointValues.forEach(points => {
        const th = document.createElement('th');
        th.textContent = points;
        headerRow.appendChild(th);
    });

    // 3. Создать тело таблицы (tbody)
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

                    // Логика "лесенки"
                    if (isFirstQuestionInRow) {
                        isFirstQuestionInRow = false;
                    } else {
                        cell.classList.add('locked');
                    }
                } else {
                    cell.className = 'placeholder-cell'; // Пустая ячейка, если вопроса с такими баллами нет
                }
            });
        }
    }
}

function handleQuestionClick(cell) {
    currentQuestionCell = cell;
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
    const points = parseInt(currentQuestionCell.dataset.points, 10);
    const answerButtons = document.querySelectorAll('#answers-container .answer-btn');

    answerButtons.forEach(btn => btn.disabled = true);

    if (selectedIndex === correctIndex) {
        answerButtons[selectedIndex].classList.add('correct');
        score += points;
        updateScore();
    } else {
        answerButtons[selectedIndex].classList.add('incorrect');
        if (correctIndex > -1) answerButtons[correctIndex].classList.add('correct');
    }

    currentQuestionCell.classList.add("disabled");

    // Разблокируем следующую ячейку в строке
    let nextCell = currentQuestionCell.nextElementSibling;
    while(nextCell && !nextCell.classList.contains('question-cell')) {
        nextCell = nextCell.nextElementSibling; // Пропускаем пустые ячейки
    }
    if (nextCell) {
        nextCell.classList.remove('locked');
    }

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
