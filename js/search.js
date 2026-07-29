/*
=========================================
 Excel AI Search Engine
=========================================
*/

function searchFormula(query) {

    query = query.toLowerCase().trim();

    let results = [];

    formulas.forEach(item => {

        let score = 0;

        // Title Match
        if (item.title.toLowerCase().includes(query)) {
            score += 100;
        }

        // Keyword Match
        if (item.keywords) {

            item.keywords.forEach(keyword => {

                if (query.includes(keyword.toLowerCase())) {
                    score += 20;
                }

            });

        }

        // Formula Match
        if (item.formula.toLowerCase().includes(query)) {
            score += 10;
        }

        if (score > 0) {

            results.push({
                ...item,
                score: score
            });

        }

    });

    // Highest score first
    results.sort((a, b) => b.score - a.score);

    return results;

}