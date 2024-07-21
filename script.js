document.addEventListener("DOMContentLoaded", () => {
    addInitialRows();
});

function addInitialRows() {
    addRows('ife', 'strengths', 5);
    addRows('ife', 'weaknesses', 5);
    addRows('efe', 'opportunities', 5);
    addRows('efe', 'threats', 5);
}

function addRows(matrixType, section, initialCount) {
    const body = document.getElementById(`${matrixType}-${section}-body`);
    for (let i = 0; i < initialCount; i++) {
        const row = document.createElement('tr');
        row.innerHTML = `
            <td><input type="text" class="factor" /></td>
            <td><input type="text" class="weight" /></td>
            <td><input type="text" class="rating" /></td>
            <td><input type="text" class="weighted-score" readonly /></td>
        `;
        body.appendChild(row);
    }
}

function addFactor(matrixType, section) {
    const body = document.getElementById(`${matrixType}-${section}-body`);
    const row = document.createElement('tr');
    row.innerHTML = `
        <td><input type="text" class="factor" /></td>
        <td><input type="text" class="weight" /></td>
        <td><input type="text" class="rating" /></td>
        <td><input type="text" class="weighted-score" readonly /></td>
    `;
    body.appendChild(row);
}

function calculateIFE() {
    calculateMatrix('ife', 'strengths');
    calculateMatrix('ife', 'weaknesses');
    updateTotal('ife');
    generateConclusion();
}

function calculateEFE() {
    calculateMatrix('efe', 'opportunities');
    calculateMatrix('efe', 'threats');
    updateTotal('efe');
    generateConclusion();
}

function calculateMatrix(matrixType, section) {
    const weights = document.querySelectorAll(`#${matrixType}-${section}-body .weight`);
    const ratings = document.querySelectorAll(`#${matrixType}-${section}-body .rating`);
    const weightedScores = document.querySelectorAll(`#${matrixType}-${section}-body .weighted-score`);
    
    weights.forEach((weight, index) => {
        const weightValue = parseFloat(weight.value) || 0;
        const ratingValue = parseFloat(ratings[index].value) || 0;
        const weightedScore = weightValue * ratingValue;
        
        weightedScores[index].value = weightedScore.toFixed(2);
    });
}

function updateTotal(matrixType) {
    const weights = document.querySelectorAll(`#${matrixType}-form .weight`);
    const weightedScores = document.querySelectorAll(`#${matrixType}-form .weighted-score`);

    let totalWeight = 0;
    let totalScore = 0;

    weights.forEach(weight => {
        totalWeight += parseFloat(weight.value) || 0;
    });

    weightedScores.forEach(score => {
        totalScore += parseFloat(score.value) || 0;
    });

    document.getElementById(`${matrixType}-total-weight`).value = totalWeight.toFixed(2);
    document.getElementById(`${matrixType}-total-score`).value = totalScore.toFixed(2);
}

function resetForms() {
    document.querySelectorAll('form input').forEach(input => input.value = '');
    document.getElementById('conclusion-text').innerText = 'Una vez calculadas las matrices, aparecerá aquí un resumen de los hallazgos y recomendaciones.';
}

function exportToExcel() {
    const ifeData = exportMatrixToArray('ife');
    const efeData = exportMatrixToArray('efe');
    const conclusionData = document.getElementById('conclusion-text').innerText;

    const ws = XLSX.utils.aoa_to_sheet([
        ['IFE Matrix - Fortalezas', '', '', '', ''],
        ['Factor', 'Peso', 'Clasificación', 'Peso Ponderado'],
        ...ifeData.strengths,
        ['Totales', document.getElementById('ife-total-weight').value, '', document.getElementById('ife-total-score').value],
        [],
        ['IFE Matrix - Debilidades', '', '', '', ''],
        ['Factor', 'Peso', 'Clasificación', 'Peso Ponderado'],
        ...ifeData.weaknesses,
        ['Totales', document.getElementById('ife-total-weight').value, '', document.getElementById('ife-total-score').value],
        [],
        ['EFE Matrix - Oportunidades', '', '', '', ''],
        ['Factor', 'Peso', 'Clasificación', 'Peso Ponderado'],
        ...efeData.opportunities,
        ['Totales', document.getElementById('efe-total-weight').value, '', document.getElementById('efe-total-score').value],
        [],
        ['EFE Matrix - Amenazas', '', '', '', ''],
        ['Factor', 'Peso', 'Clasificación', 'Peso Ponderado'],
        ...efeData.threats,
        ['Totales', document.getElementById('efe-total-weight').value, '', document.getElementById('efe-total-score').value],
        [],
        ['Conclusión'],
        [conclusionData]
    ]);

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Matrices');

    XLSX.writeFile(wb, 'IFE_EFE_Matrices.xlsx');
}

function exportMatrixToArray(matrixType) {
    const strengths = [];
    const weaknesses = [];
    const opportunities = [];
    const threats = [];

    document.querySelectorAll(`#${matrixType}-strengths-body tr`).forEach(row => {
        const factor = row.querySelector('.factor').value;
        const weight = row.querySelector('.weight').value;
        const rating = row.querySelector('.rating').value;
        const weightedScore = row.querySelector('.weighted-score').value;
        if (factor && weight && rating && weightedScore) {
            strengths.push([factor, weight, rating, weightedScore]);
        }
    });

    document.querySelectorAll(`#${matrixType}-weaknesses-body tr`).forEach(row => {
        const factor = row.querySelector('.factor').value;
        const weight = row.querySelector('.weight').value;
        const rating = row.querySelector('.rating').value;
        const weightedScore = row.querySelector('.weighted-score').value;
        if (factor && weight && rating && weightedScore) {
            weaknesses.push([factor, weight, rating, weightedScore]);
        }
    });

    document.querySelectorAll(`#${matrixType}-opportunities-body tr`).forEach(row => {
        const factor = row.querySelector('.factor').value;
        const weight = row.querySelector('.weight').value;
        const rating = row.querySelector('.rating').value;
        const weightedScore = row.querySelector('.weighted-score').value;
        if (factor && weight && rating && weightedScore) {
            opportunities.push([factor, weight, rating, weightedScore]);
        }
    });

    document.querySelectorAll(`#${matrixType}-threats-body tr`).forEach(row => {
        const factor = row.querySelector('.factor').value;
        const weight = row.querySelector('.weight').value;
        const rating = row.querySelector('.rating').value;
        const weightedScore = row.querySelector('.weighted-score').value;
        if (factor && weight && rating && weightedScore) {
            threats.push([factor, weight, rating, weightedScore]);
        }
    });

    return { strengths, weaknesses, opportunities, threats };
}

function generateConclusion() {
    const ifeTotalScore = parseFloat(document.getElementById('ife-total-score').value) || 0;
    const efeTotalScore = parseFloat(document.getElementById('efe-total-score').value) || 0;

    let conclusionText = `La puntuación total de la matriz IFE es ${ifeTotalScore.toFixed(2)} y la de la matriz EFE es ${efeTotalScore.toFixed(2)}.\n\n`;
    
    conclusionText += "Factores IFE:\n";
    conclusionText += generateFactorsConclusion('ife', 'strengths', "Fortalezas");
    conclusionText += generateFactorsConclusion('ife', 'weaknesses', "Debilidades");

    conclusionText += "\nFactores EFE:\n";
    conclusionText += generateFactorsConclusion('efe', 'opportunities', "Oportunidades");
        conclusionText += generateFactorsConclusion('efe', 'threats', "Amenazas");

    if (ifeTotalScore < 2.50) {
        conclusionText += "\nLa puntuación IFE es menor que el umbral de 2.50, lo que indica que hay áreas internas que necesitan mejorar.\n";
    } else {
        conclusionText += "\nLa puntuación IFE está por encima del umbral de 2.50, indicando una fuerte posición interna.\n";
    }

    if (efeTotalScore < 2.50) {
        conclusionText += "La puntuación EFE es menor que el umbral de 2.50, lo que indica que hay desafíos externos significativos.\n";
    } else {
        conclusionText += "La puntuación EFE está por encima del umbral de 2.50, indicando una fuerte capacidad para aprovechar oportunidades y mitigar amenazas externas.\n";
    }

    document.getElementById('conclusion-text').innerText = conclusionText;
}

function generateFactorsConclusion(matrixType, section, sectionTitle) {
    const factors = document.querySelectorAll(`#${matrixType}-${section}-body .factor`);
    const weights = document.querySelectorAll(`#${matrixType}-${section}-body .weight`);
    const ratings = document.querySelectorAll(`#${matrixType}-${section}-body .rating`);
    const weightedScores = document.querySelectorAll(`#${matrixType}-${section}-body .weighted-score`);

    let conclusion = `${sectionTitle}:\n`;

    factors.forEach((factor, index) => {
        const factorValue = factor.value;
        const weightValue = parseFloat(weights[index].value) || 0;
        const ratingValue = parseFloat(ratings[index].value) || 0;
        const weightedScoreValue = parseFloat(weightedScores[index].value) || 0;

        if (factorValue) {
            conclusion += `- ${factorValue}: Peso ${weightValue.toFixed(2)}, Clasificación ${ratingValue.toFixed(2)}, Peso Ponderado ${weightedScoreValue.toFixed(2)}\n`;
        }
    });

    return conclusion + "\n";
}

