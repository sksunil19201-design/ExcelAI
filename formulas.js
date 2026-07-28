// ==========================================================
// ExcelAI Formula Database
// Powered By Sunil.K
// ==========================================================

const formulas = [

    {
        title:"SUM",
        keywords:[
            "sum",
            "total"
        ],
        formula:"=SUM(A2:A10)"
    },

    {
        title:"XLOOKUP",
        keywords:[
            "lookup",
            "find value"
        ],
        formula:"=XLOOKUP(A2,D:D,E:E)"
    }

    // Your Part 1-60 formulas continue here

];


// ==========================================================
// Export
// ==========================================================

// Make formulas available for AI engine
window.formulas = formulas;

console.log("Formula Database Loaded:", formulas.length);