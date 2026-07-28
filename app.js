/*
=========================================
 Excel AI
 app.js
 Version : 1.0
=========================================
*/

document.addEventListener("DOMContentLoaded", () => {

    console.log("✅ App Started");

    // Elements
    const prompt = document.getElementById("prompt");
    const generateBtn = document.getElementById("generateBtn");

    // Generate Button
    generateBtn.addEventListener("click", generateAnswer);

    // Press Enter
    prompt.addEventListener("keydown", function (e) {

        if (e.key === "Enter" && !e.shiftKey) {

            e.preventDefault();

            generateAnswer();

        }

    });

});


/*=========================================
 Generate Answer
=========================================*/

function generateAnswer() {

    const question = document
        .getElementById("prompt")
        .value
        .trim();

    if (question === "") {

        alert("Please enter your Excel question.");

        return;

    }

    const result = searchFormula(question);

    if (!result) {

        showNotFound();

        return;

    }

    fillResult(result);

}


/*=========================================
 Fill Result
=========================================*/

function fillResult(result) {

    // Formula
    document.getElementById("formulaBox").textContent =
        result.formula;

    // Explanation
    document.getElementById("explanation").textContent =
        result.explanation;

    // Business
    document.getElementById("businessCase").textContent =
        result.business;

    // Output
    document.getElementById("outputValue").textContent =
        result.output;

    // Alternative
    document.getElementById("alternativeFormula").textContent =
        result.alternative;

    // Breakdown
    const breakdown = document.getElementById("breakdown");

    breakdown.innerHTML = "";

    result.breakdown.forEach(item => {

        breakdown.innerHTML += `<li>${item}</li>`;

    });

    // Tips
    const tips = document.getElementById("tipsList");

    tips.innerHTML = "";

    result.tips.forEach(item => {

        tips.innerHTML += `<li>${item}</li>`;

    });

    // Errors
    const errors = document.getElementById("errorList");

    errors.innerHTML = "";

    result.errors.forEach(item => {

        errors.innerHTML += `<li>${item}</li>`;

    });

    // Table
    renderTable(result.dataset);

    // History
    addHistory(result.formula);

}


/*=========================================
 Search History
=========================================*/

function addHistory(text) {

    const history = document.getElementById("historyList");

    if (!history) return;

    const li = document.createElement("li");

    li.innerText = text;

    history.prepend(li);

}


/*=========================================
 Formula Not Found
=========================================*/

function showNotFound() {

    document.getElementById("formulaBox").textContent =
        "❌ Formula not found";

    document.getElementById("explanation").textContent =
        "Try another Excel question.";

    document.getElementById("businessCase").textContent =
        "-";

    document.getElementById("outputValue").textContent =
        "-";

    document.getElementById("alternativeFormula").textContent =
        "-";

    document.getElementById("tipsList").innerHTML = "";

    document.getElementById("errorList").innerHTML = "";

    document.getElementById("breakdown").innerHTML = "";

    document.getElementById("tableContainer").innerHTML = "";

}