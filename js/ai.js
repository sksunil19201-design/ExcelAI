/*
==========================================
 Excel AI
 ai.js
 Version : 1.0
==========================================
*/

const AI_DATABASE = [

// ==========================================================
// VLOOKUP
// ==========================================================
{
    title: "VLOOKUP",

    keywords: [
        "vlookup","lookup","vertical lookup","find value","search value",
        "employee","employee id","salary","department","customer",
        "customer id","product","product code","invoice","price",
        "fetch data","retrieve data","return value","reference table",
        "master table","database lookup","exact match","table lookup",
        "excel lookup","find employee","employee salary","bring data",
        "id lookup","search employee","lookup employee","record search",
        "data mapping","merge tables","match id","find record","old lookup"
    ],

    aliases: [
        "employee salary lookup",
        "customer lookup",
        "product lookup",
        "lookup by id",
        "search employee",
        "find employee salary",
        "retrieve information",
        "bring value from table",
        "table reference",
        "excel database lookup"
    ],

    questions: [
        "How to find employee salary?",
        "How to lookup customer details?",
        "How to search using employee ID?",
        "How to fetch value from another table?",
        "How to match two tables?",
        "How to return department name?",
        "How to lookup product price?",
        "How to search invoice details?",
        "How to retrieve data using ID?",
        "How to perform vertical lookup?"
    ],

    dataset: "employees",

    category: "Lookup",

    formula: "=VLOOKUP(H2,A2:E100,4,FALSE)",

    syntax: "VLOOKUP(lookup_value,table_array,col_index,FALSE)",

    explanation: "Searches the first column of a table and returns a value from another column.",

    breakdown: [
        "H2 = Employee ID",
        "A:E = Employee Table",
        "4 = Salary Column",
        "FALSE = Exact Match"
    ],

    output: "$65,000",

    business: "Retrieve employee salary, customer details, product price, or department instantly.",

    tips: [
        "Lookup value must exist in first column.",
        "Use FALSE for exact match.",
        "Keep lookup table sorted only if using TRUE."
    ],

    errors: [
        "#N/A",
        "Wrong column number",
        "Lookup value not found",
        "Leading/trailing spaces"
    ],

    alternative: "=XLOOKUP(H2,A2:A100,D2:D100)"
},

// ==========================================================
// XLOOKUP
// ==========================================================
{
    title: "XLOOKUP",

    keywords: [
        "xlookup","lookup","modern lookup","excel 365","excel 2021",
        "employee","salary","customer","product","invoice",
        "search value","find value","return value","left lookup",
        "right lookup","lookup id","employee id","customer id",
        "retrieve data","fetch data","exact match","approximate match",
        "lookup table","dynamic lookup","search employee",
        "employee salary","product lookup","replacement for vlookup",
        "best lookup","advanced lookup","modern excel","data mapping",
        "return array","lookup record","flexible lookup","lookup formula"
    ],

    aliases: [
        "modern vlookup",
        "advanced lookup",
        "left lookup",
        "right lookup",
        "employee lookup",
        "customer search",
        "product finder",
        "dynamic lookup",
        "new lookup function",
        "excel lookup"
    ],

    questions: [
        "How to replace VLOOKUP?",
        "How to lookup from left to right?",
        "How to return multiple values?",
        "How to search employee salary?",
        "How to find customer information?",
        "How to lookup product price?",
        "How to perform flexible lookup?",
        "How to search any column?",
        "How to lookup using Excel 365?",
        "How to return exact match?"
    ],

    dataset: "employees",

    category: "Lookup",

    formula: "=XLOOKUP(H2,A2:A100,D2:D100)",

    syntax: "XLOOKUP(value,lookup_array,return_array)",

    explanation: "Searches any column and returns values without column index limitations.",

    breakdown: [
        "H2 = Employee ID",
        "A:A = Lookup Column",
        "D:D = Salary Column"
    ],

    output: "$65,000",

    business: "Modern replacement for VLOOKUP with greater flexibility and accuracy.",

    tips: [
        "Works left-to-right and right-to-left.",
        "Supports exact match by default.",
        "Preferred for Excel 365."
    ],

    errors: [
        "#N/A",
        "Lookup value missing",
        "Incorrect return array"
    ],

    alternative: "=INDEX(D2:D100,MATCH(H2,A2:A100,0))"
},

// ==========================================================
// HLOOKUP
// ==========================================================
{
    title: "HLOOKUP",

    keywords: [
        "hlookup","horizontal lookup","lookup","find value","search value",
        "row lookup","header lookup","horizontal search","retrieve value",
        "sales","months","quarter","budget","forecast","excel lookup",
        "lookup row","return row value","find month","lookup month",
        "horizontal table","table lookup","financial report",
        "cross sheet lookup","row match","header search","find header",
        "horizontal database","lookup formula","spreadsheet lookup",
        "top row lookup","row reference","lookup across columns"
    ],

    aliases: [
        "horizontal search",
        "row lookup",
        "month lookup",
        "header lookup",
        "find value across columns",
        "horizontal table lookup",
        "lookup by header",
        "sales lookup",
        "budget lookup",
        "forecast lookup"
    ],

    questions: [
        "How to lookup data horizontally?",
        "How to search month values?",
        "How to retrieve quarterly sales?",
        "How to find data by header?",
        "How to lookup row values?",
        "How to search top row?",
        "How to return budget values?",
        "How to lookup forecast?",
        "How to perform horizontal lookup?",
        "How to search across columns?"
    ],

    dataset: "sales",

    category: "Lookup",

    formula: "=HLOOKUP(B1,A1:G5,3,FALSE)",

    syntax: "HLOOKUP(value,table,row_index,FALSE)",

    explanation: "Searches the first row of a table and returns a value from another row.",

    breakdown: [
        "B1 = Lookup Header",
        "A:G = Data Table",
        "3 = Return Row",
        "FALSE = Exact Match"
    ],

    output: "$120,000",

    business: "Retrieve monthly sales, budget, or forecast values arranged horizontally.",

    tips: [
        "Header must be in the first row.",
        "Use FALSE for exact matches.",
        "Ideal for horizontal datasets."
    ],

    errors: [
        "#N/A",
        "Wrong row index",
        "Header not found"
    ],

    alternative: "=XLOOKUP(B1,A1:G1,A3:G3)"
},


    // ==========================================================
// LOOKUP
// ==========================================================
{
    title: "LOOKUP",

    keywords: [
        "lookup","vector lookup","array lookup","search value","find value",
        "lookup function","approximate lookup","return value","find record",
        "employee","customer","product","sales","price","invoice",
        "code lookup","lookup table","excel lookup","retrieve data",
        "fetch value","database lookup","master table","search record",
        "match value","lookup code","product code","customer id",
        "employee id","find price","find customer","find employee",
        "legacy lookup","old excel","array search","vector search",
        "reference data","data lookup","quick lookup","search code"
    ],

    aliases: [
        "vector search",
        "legacy lookup",
        "array lookup",
        "search table",
        "find from list",
        "retrieve value",
        "lookup by code",
        "old lookup function",
        "excel lookup",
        "basic lookup"
    ],

    questions: [
        "How to use LOOKUP function?",
        "How to search a value in Excel?",
        "How to retrieve product price?",
        "How to lookup customer name?",
        "How to return data from a list?",
        "How to search using product code?",
        "How to find employee details?",
        "How to lookup values from an array?",
        "How to perform approximate lookup?",
        "How to use legacy lookup?"
    ],

    dataset:"products",

    category:"Lookup",

    formula:"=LOOKUP(H2,A2:A100,C2:C100)",

    syntax:"LOOKUP(lookup_value,lookup_vector,result_vector)",

    explanation:"Searches a single row or column and returns the corresponding value from another row or column.",

    breakdown:[
        "H2 = Lookup Value",
        "A:A = Search Column",
        "C:C = Return Column"
    ],

    output:"Laptop",

    business:"Retrieve product names, prices, customer information, or employee records from simple lookup tables.",

    tips:[
        "LOOKUP performs approximate matching.",
        "Sort lookup values in ascending order.",
        "Prefer XLOOKUP for modern Excel."
    ],

    errors:[
        "#N/A",
        "Lookup vector not sorted",
        "Wrong lookup range"
    ],

    alternative:"=XLOOKUP(H2,A2:A100,C2:C100)"
},

// ==========================================================
// XMATCH
// ==========================================================
{
    title:"XMATCH",

    keywords:[
        "xmatch","match","modern match","position","find position",
        "lookup position","row number","column number","search index",
        "exact match","approximate match","reverse search","binary search",
        "employee","customer","product","invoice","employee id",
        "customer id","product code","search record","find record",
        "return position","index position","excel 365","dynamic lookup",
        "lookup helper","advanced match","replacement for match",
        "find row","find column","lookup index","search employee",
        "find customer","search product","position lookup","array position",
        "modern excel","match function"
    ],

    aliases:[
        "modern MATCH",
        "position finder",
        "row finder",
        "column finder",
        "advanced match",
        "search index",
        "lookup position",
        "excel position",
        "dynamic match",
        "replacement for MATCH"
    ],

    questions:[
        "How to find row number?",
        "How to replace MATCH?",
        "How to find product position?",
        "How to return row index?",
        "How to search employee position?",
        "How to find column number?",
        "How to perform reverse lookup?",
        "How to use XMATCH?",
        "How to return record position?",
        "How to search dynamically?"
    ],

    dataset:"employees",

    category:"Lookup",

    formula:"=XMATCH(H2,A2:A100)",

    syntax:"XMATCH(lookup_value,lookup_array)",

    explanation:"Returns the relative position of a lookup value with support for exact, approximate, reverse, and binary searches.",

    breakdown:[
        "H2 = Lookup Value",
        "A:A = Lookup Array",
        "Returns Position Number"
    ],

    output:"15",

    business:"Locate the row position of employees, customers, invoices, or products before retrieving related data.",

    tips:[
        "Supports reverse search.",
        "Works well with INDEX.",
        "Preferred over MATCH in Microsoft 365."
    ],

    errors:[
        "#N/A",
        "Lookup value not found",
        "Invalid search mode"
    ],

    alternative:"=MATCH(H2,A2:A100,0)"
},

// ==========================================================
// INDEX
// ==========================================================
{
    title:"INDEX",

    keywords:[
        "index","return value","retrieve value","find data","lookup",
        "advanced lookup","table lookup","array lookup","row lookup",
        "column lookup","employee","salary","customer","product",
        "invoice","sales","database","reference","return cell value",
        "extract data","fetch data","find position","dynamic lookup",
        "excel formula","data retrieval","reporting","dashboard",
        "lookup table","master data","search record","find information",
        "return result","cell reference","range lookup","data analysis",
        "business report","excel analyst"
    ],

    aliases:[
        "return value from table",
        "advanced lookup formula",
        "extract value",
        "retrieve cell data",
        "table value finder",
        "dynamic reference",
        "lookup alternative",
        "data extraction formula",
        "find table value",
        "excel reference function"
    ],

    questions:[
        "How to return value from a table?",
        "How to extract data from Excel range?",
        "How to find employee salary?",
        "How to retrieve value using row number?",
        "How to return data by position?",
        "How to create dynamic lookup?",
        "How to get value from another table?",
        "How to use INDEX formula?",
        "How to find data using row and column?",
        "How to create advanced lookup?"
    ],

    dataset:"employees",

    category:"Lookup",

    formula:"=INDEX(D2:D100,MATCH(H2,A2:A100,0))",

    syntax:"INDEX(array,row_num,[column_num])",

    explanation:"Returns a value from a specific position inside a range or array.",

    breakdown:[
        "D2:D100 = Return Range",
        "MATCH finds employee position",
        "INDEX returns matching salary"
    ],

    output:"$65,000",

    business:"Retrieve employee salary, customer information, or product details from large datasets.",

    tips:[
        "Works well with MATCH.",
        "More flexible than VLOOKUP.",
        "Can lookup left or right."
    ],

    errors:[
        "#REF!",
        "#VALUE!",
        "Invalid row number",
        "Wrong range selection"
    ],

    alternative:"=XLOOKUP(H2,A2:A100,D2:D100)"
},


// ==========================================================
// MATCH
// ==========================================================
{
    title:"MATCH",

    keywords:[
        "match","find position","position lookup","row number",
        "column number","search position","lookup position",
        "find location","index number","employee id","customer id",
        "product code","invoice number","search value",
        "exact match","approximate match","lookup helper",
        "index match","advanced lookup","data search",
        "find record","find row","find column","excel lookup",
        "table search","dynamic search","data mapping",
        "reference position","array search","record location",
        "formula helper","excel analysis","report automation",
        "find matching value"
    ],

    aliases:[
        "position finder",
        "row locator",
        "column locator",
        "search index",
        "lookup helper",
        "find matching row",
        "value position",
        "excel search function",
        "record finder",
        "table position finder"
    ],

    questions:[
        "How to find row number in Excel?",
        "How to locate a value?",
        "How to find employee position?",
        "How to find matching record?",
        "How to search exact value?",
        "How to return position number?",
        "How to use MATCH function?",
        "How to find column position?",
        "How to combine INDEX and MATCH?",
        "How to search data dynamically?"
    ],

    dataset:"employees",

    category:"Lookup",

    formula:"=MATCH(H2,A2:A100,0)",

    syntax:"MATCH(lookup_value,lookup_array,match_type)",

    explanation:"Returns the relative position of a value in a row or column.",

    breakdown:[
        "H2 = Value to search",
        "A2:A100 = Search Range",
        "0 = Exact Match"
    ],

    output:"15",

    business:"Find employee position, customer record location, or product row number.",

    tips:[
        "Use 0 for exact matching.",
        "Commonly used with INDEX.",
        "Works vertically and horizontally."
    ],

    errors:[
        "#N/A",
        "Value not found",
        "Incorrect match type"
    ],

    alternative:"=XMATCH(H2,A2:A100)"
},

// ==========================================================
// INDEX + MATCH
// ==========================================================
{
    title:"INDEX MATCH",

    keywords:[
        "index match","index and match","advanced lookup",
        "vlookup alternative","lookup","employee salary",
        "find employee","employee id","customer lookup",
        "product lookup","dynamic lookup","left lookup",
        "right lookup","two way lookup","flexible lookup",
        "table search","retrieve data","fetch value",
        "return value","find record","database lookup",
        "master data lookup","excel analyst","data analyst",
        "large dataset lookup","complex lookup","exact match",
        "lookup formula","replace vlookup","advanced excel",
        "report automation","dashboard formula","data retrieval",
        "search table","match position","dynamic reference"
    ],

    aliases:[
        "advanced vlookup",
        "vlookup replacement",
        "flexible lookup formula",
        "two way lookup",
        "employee data finder",
        "dynamic table lookup",
        "professional lookup",
        "analyst lookup formula",
        "left side lookup",
        "large data lookup"
    ],

    questions:[
        "How to replace VLOOKUP?",
        "How to lookup left side data?",
        "How to find employee salary using ID?",
        "How to perform advanced lookup?",
        "How to lookup from multiple columns?",
        "How to create flexible lookup?",
        "How to search large datasets?",
        "How to combine INDEX and MATCH?",
        "How to return matching value?",
        "How to build analyst level lookup?"
    ],

    dataset:"employees",

    category:"Lookup",

    formula:"=INDEX(D2:D100,MATCH(H2,A2:A100,0))",

    syntax:"INDEX(return_range,MATCH(lookup_value,lookup_range,0))",

    explanation:"Combines INDEX and MATCH to perform flexible lookups without VLOOKUP limitations.",

    breakdown:[
        "MATCH finds employee position",
        "INDEX returns corresponding salary",
        "0 ensures exact match"
    ],

    output:"$65,000",

    business:"Used by analysts to retrieve information from large business datasets and reports.",

    tips:[
        "Can lookup left or right.",
        "Works in older Excel versions.",
        "More flexible than VLOOKUP."
    ],

    errors:[
        "#N/A",
        "#REF!",
        "Incorrect lookup range",
        "Mismatch in ranges"
    ],

    alternative:"=XLOOKUP(H2,A2:A100,D2:D100)"
},

// ==========================================================
// CHOOSE
// ==========================================================
{
    title:"CHOOSE",

    keywords:[
        "choose","select value","pick value","return option",
        "multiple choice","index selection","conditional selection",
        "employee grade","category","rating","status",
        "priority","level","mapping","lookup alternative",
        "value selection","data mapping","assign category",
        "return based on number","dynamic selection",
        "excel decision","formula selection","option finder",
        "business classification","rank mapping","score mapping",
        "grade calculation","commission level","bonus calculation",
        "sales category","performance rating","choice function"
    ],

    aliases:[
        "select from list",
        "return selected value",
        "option selector",
        "category mapping",
        "value picker",
        "formula based selection",
        "rating mapper",
        "grade generator",
        "decision formula",
        "choice formula"
    ],

    questions:[
        "How to select value based on number?",
        "How to create grade formula?",
        "How to assign categories?",
        "How to return different results?",
        "How to create option selection?",
        "How to map scores to levels?",
        "How to calculate performance rating?",
        "How to choose from multiple values?",
        "How to create simple decision logic?",
        "How to return matching option?"
    ],

    dataset:"performance",

    category:"Lookup",

    formula:"=CHOOSE(B2,\"Low\",\"Medium\",\"High\")",

    syntax:"CHOOSE(index_num,value1,value2,value3...)",

    explanation:"Returns a value from a list based on an index number.",

    breakdown:[
        "B2 = Selection Number",
        "Value1 = First Option",
        "Value2 = Second Option",
        "Value3 = Third Option"
    ],

    output:"High",

    business:"Create rating systems, employee grades, categories, and business classifications.",

    tips:[
        "Index number should be between available options.",
        "Useful for small mapping tasks.",
        "Combine with MATCH for dynamic selection."
    ],

    errors:[
        "#VALUE!",
        "Invalid index number",
        "Missing option"
    ],

    alternative:"=IFS(B2=1,\"Low\",B2=2,\"Medium\",B2=3,\"High\")"
},


// ==========================================================
// OFFSET
// ==========================================================
{
    title:"OFFSET",

    keywords:[
        "offset","dynamic range","dynamic reference",
        "moving range","variable range","dashboard range",
        "chart range","dynamic chart","excel reference",
        "return cell reference","shift reference",
        "data range","automatic update","rolling data",
        "monthly report","dashboard automation",
        "dynamic formula","range selection","table expansion",
        "report automation","financial model","forecast",
        "sales report","dynamic table","change reference",
        "cell movement","reference formula","advanced excel",
        "interactive dashboard","data analysis",
        "auto expanding range","live report"
    ],

    aliases:[
        "dynamic range formula",
        "moving reference",
        "automatic range",
        "dynamic chart formula",
        "flexible reference",
        "changing range",
        "dashboard formula",
        "rolling range",
        "live range",
        "advanced reference"
    ],

    questions:[
        "How to create dynamic range?",
        "How to make chart update automatically?",
        "How to create moving reference?",
        "How to create rolling reports?",
        "How to expand data range automatically?",
        "How to use OFFSET formula?",
        "How to create dashboard ranges?",
        "How to make formulas dynamic?",
        "How to create live charts?",
        "How to reference changing cells?"
    ],

    dataset:"sales",

    category:"Lookup",

    formula:"=OFFSET(A1,0,2,10,1)",

    syntax:"OFFSET(reference,rows,cols,height,width)",

    explanation:"Returns a reference shifted from a starting cell by specified rows and columns.",

    breakdown:[
        "A1 = Starting Cell",
        "0 = Move Rows",
        "2 = Move Columns",
        "10 = Height",
        "1 = Width"
    ],

    output:"Dynamic Range",

    business:"Used for dynamic dashboards, charts, reports, and automatically expanding ranges.",

    tips:[
        "Avoid excessive use in large files.",
        "Useful for dynamic charts.",
        "Can create named dynamic ranges."
    ],

    errors:[
        "#REF!",
        "Range outside worksheet",
        "Incorrect height or width"
    ],

    alternative:"=INDEX(A:A,1):INDEX(A:A,COUNTA(A:A))"
},

// ==========================================================
// INDIRECT
// ==========================================================
{
    title:"INDIRECT",

    keywords:[
        "indirect","dynamic reference","cell reference",
        "dynamic formula","reference cell","text reference",
        "convert text to reference","dynamic lookup",
        "variable cell","variable range","change reference",
        "dynamic worksheet","sheet reference","multiple sheets",
        "dashboard formula","report automation","advanced excel",
        "flexible reference","formula builder","reference generator",
        "dynamic table","dynamic report","monthly report",
        "financial model","scenario analysis","data validation",
        "user input reference","automatic formula",
        "create reference from text","lookup reference",
        "excel automation","cell address","range reference",
        "sheet automation","interactive dashboard"
    ],

    aliases:[
        "dynamic cell reference",
        "text based reference",
        "reference generator",
        "variable lookup",
        "formula from text",
        "dynamic sheet lookup",
        "flexible cell reference",
        "automatic reference",
        "advanced reference formula",
        "dynamic reporting formula"
    ],

    questions:[
        "How to create dynamic cell reference?",
        "How to reference another sheet dynamically?",
        "How to convert text into cell reference?",
        "How to create flexible formulas?",
        "How to build dynamic reports?",
        "How to change reference automatically?",
        "How to create variable range?",
        "How to use INDIRECT function?",
        "How to link multiple sheets dynamically?",
        "How to create dashboard references?"
    ],

    dataset:"reports",

    category:"Lookup",

    formula:'=INDIRECT("D"&A2)',

    syntax:"INDIRECT(ref_text,[a1])",

    explanation:"Returns a cell or range reference created from a text string.",

    breakdown:[
        "A2 = Row Number",
        "\"D\" = Column Reference",
        "INDIRECT converts text into cell reference"
    ],

    output:"Employee Salary Cell",

    business:"Create dynamic dashboards, reports, and formulas where references change based on user input.",

    tips:[
        "Useful for dynamic sheet references.",
        "Avoid excessive use in very large files.",
        "Works well with dropdown selections."
    ],

    errors:[
        "#REF!",
        "Invalid reference text",
        "Deleted sheet reference"
    ],

    alternative:"=INDEX(D:D,A2)"
},


// ==========================================================
// ADDRESS
// ==========================================================
{
    title:"ADDRESS",

    keywords:[
        "address","cell address","cell reference",
        "create cell address","row column reference",
        "dynamic address","reference generator",
        "return address","find location","cell location",
        "excel reference","formula builder","dynamic formula",
        "row number","column number","cell position",
        "coordinate","range address","table location",
        "dashboard formula","automation","reporting",
        "dynamic lookup","create reference",
        "combine row column","excel automation",
        "cell identifier","data location",
        "reference text","formula generation",
        "advanced excel","dynamic reports"
    ],

    aliases:[
        "cell locator",
        "reference creator",
        "address generator",
        "cell coordinate",
        "location finder",
        "dynamic reference maker",
        "excel cell position",
        "formula reference",
        "cell pointer",
        "reference builder"
    ],

    questions:[
        "How to get cell address?",
        "How to create reference from row and column?",
        "How to find cell location?",
        "How to generate Excel address?",
        "How to create dynamic reference?",
        "How to combine row and column number?",
        "How to build formula references?",
        "How to identify cell position?",
        "How to automate cell references?",
        "How to use ADDRESS function?"
    ],

    dataset:"worksheet",

    category:"Lookup",

    formula:"=ADDRESS(5,3)",

    syntax:"ADDRESS(row_num,column_num,[abs_num])",

    explanation:"Creates a cell address based on row and column numbers.",

    breakdown:[
        "5 = Row Number",
        "3 = Column Number",
        "Returns cell reference"
    ],

    output:"$C$5",

    business:"Generate dynamic cell references for automation, reporting, and advanced formulas.",

    tips:[
        "Combine with INDIRECT for dynamic references.",
        "Use ROW and COLUMN for automation.",
        "Can return absolute or relative references."
    ],

    errors:[
        "#VALUE!",
        "Invalid row number",
        "Invalid column number"
    ],

    alternative:"=INDIRECT(ADDRESS(5,3))"
},

// ==========================================================
// ROW
// ==========================================================
{
    title:"ROW",

    keywords:[
        "row","row number","get row number","find row",
        "current row","return row","row reference",
        "cell position","record number","serial number",
        "auto numbering","dynamic numbering","index number",
        "data row","excel position","table row",
        "running number","sequence number","formula row",
        "dynamic formula","automation","reporting",
        "dashboard","data analysis","reference row",
        "find record position","row calculation",
        "generate serial number","excel automation",
        "worksheet row","cell location","row identifier"
    ],

    aliases:[
        "row finder",
        "row locator",
        "record number generator",
        "automatic numbering",
        "row position",
        "current row formula",
        "dynamic row reference",
        "excel row counter",
        "row identifier",
        "position formula"
    ],

    questions:[
        "How to get row number?",
        "How to create automatic serial numbers?",
        "How to find current row?",
        "How to generate row reference?",
        "How to number records automatically?",
        "How to identify row position?",
        "How to use ROW function?",
        "How to create dynamic numbering?",
        "How to find employee row number?",
        "How to calculate record position?"
    ],

    dataset:"employees",

    category:"Reference",

    formula:"=ROW(A5)",

    syntax:"ROW([reference])",

    explanation:"Returns the row number of a specified cell reference.",

    breakdown:[
        "A5 = Referenced Cell",
        "Returns row position number"
    ],

    output:"5",

    business:"Create automatic numbering, identify records, and build dynamic reports.",

    tips:[
        "Combine with ROWS for numbering.",
        "Useful when copying formulas down.",
        "Works without any reference."
    ],

    errors:[
        "#VALUE!",
        "Invalid reference"
    ],

    alternative:"=ROWS($A$2:A2)"
},


// ==========================================================
// COLUMN
// ==========================================================
{
    title:"COLUMN",

    keywords:[
        "column","column number","get column number",
        "find column","current column","return column",
        "column reference","cell position","header position",
        "table column","data column","excel position",
        "dynamic formula","column calculation",
        "automation","dashboard","reporting",
        "find field position","database column",
        "column index","reference column",
        "data analysis","excel automation",
        "worksheet column","column identifier",
        "formula position","dynamic lookup",
        "find column location","table structure",
        "range analysis"
    ],

    aliases:[
        "column finder",
        "column locator",
        "field position",
        "column identifier",
        "header position",
        "excel column number",
        "column reference formula",
        "dynamic column finder",
        "table field locator",
        "position formula"
    ],

    questions:[
        "How to get column number?",
        "How to find current column?",
        "How to identify column position?",
        "How to return column index?",
        "How to find header location?",
        "How to create dynamic columns?",
        "How to use COLUMN function?",
        "How to calculate column position?",
        "How to automate column reference?",
        "How to find Excel field number?"
    ],

    dataset:"sales",

    category:"Reference",

    formula:"=COLUMN(C5)",

    syntax:"COLUMN([reference])",

    explanation:"Returns the column number of a specified cell reference.",

    breakdown:[
        "C5 = Referenced Cell",
        "Returns column position"
    ],

    output:"3",

    business:"Used in dynamic reports, dashboards, and formulas requiring automatic column detection.",

    tips:[
        "Combine with INDEX for dynamic lookup.",
        "Useful for horizontal calculations.",
        "Can generate column sequences."
    ],

    errors:[
        "#VALUE!",
        "Invalid reference"
    ],

    alternative:"=COLUMNS(A:C)"
},


// ==========================================================
// ROWS
// ==========================================================
{
    title:"ROWS",

    keywords:[
        "rows","count rows","number of rows",
        "count records","table size","range count",
        "data count","record count","employee count",
        "customer count","dataset size","range length",
        "number of records","dynamic count",
        "excel rows","table rows","report count",
        "data analysis","dashboard metric",
        "count entries","calculate records",
        "list count","range calculation",
        "automatic count","dynamic range",
        "database rows","excel automation",
        "record calculation","data validation"
    ],

    aliases:[
        "row counter",
        "record counter",
        "range row count",
        "dataset counter",
        "table size calculator",
        "number of records",
        "list counter",
        "data volume calculator",
        "range size",
        "row measurement"
    ],

    questions:[
        "How to count rows in Excel?",
        "How many records are in a table?",
        "How to calculate dataset size?",
        "How to count employee records?",
        "How to find number of rows?",
        "How to measure range length?",
        "How to use ROWS function?",
        "How to count table entries?",
        "How to create dynamic row count?",
        "How to count data records?"
    ],

    dataset:"employees",

    category:"Reference",

    formula:"=ROWS(A2:A100)",

    syntax:"ROWS(array)",

    explanation:"Returns the number of rows in a specified range.",

    breakdown:[
        "A2:A100 = Selected Range",
        "Counts available rows"
    ],

    output:"99",

    business:"Measure dataset size, count records, and build dynamic reporting logic.",

    tips:[
        "Works with tables and arrays.",
        "Useful for dynamic ranges.",
        "Combine with INDEX for advanced formulas."
    ],

    errors:[
        "#VALUE!",
        "Invalid range"
    ],

    alternative:"=COUNTA(A2:A100)"
},


// ==========================================================
// COLUMNS
// ==========================================================
{
    title:"COLUMNS",

    keywords:[
        "columns","count columns","number of columns",
        "table width","range width","field count",
        "database columns","header count","dataset structure",
        "excel columns","count fields","range calculation",
        "data analysis","reporting","dashboard",
        "table size","dynamic range","column count",
        "number of fields","data structure",
        "excel automation","range measurement",
        "horizontal data","header analysis",
        "column calculation","dynamic formula",
        "table analysis","worksheet structure",
        "data model","excel reporting"
    ],

    aliases:[
        "column counter",
        "field counter",
        "table width calculator",
        "range column count",
        "header counter",
        "dataset structure checker",
        "column measurement",
        "excel field count",
        "range width",
        "table size formula"
    ],

    questions:[
        "How to count columns in Excel?",
        "How many fields are in a table?",
        "How to find table width?",
        "How to count headers?",
        "How to calculate column count?",
        "How to use COLUMNS function?",
        "How to find number of fields?",
        "How to analyze dataset structure?",
        "How to create dynamic column count?",
        "How to count Excel columns?"
    ],

    dataset:"sales",

    category:"Reference",

    formula:"=COLUMNS(A1:F1)",

    syntax:"COLUMNS(array)",

    explanation:"Returns the number of columns in a specified range.",

    breakdown:[
        "A1:F1 = Selected Range",
        "Counts columns available"
    ],

    output:"6",

    business:"Analyze datasets, build dynamic reports, and automate Excel models.",

    tips:[
        "Useful with dynamic arrays.",
        "Works with tables.",
        "Combine with INDEX for dynamic selection."
    ],

    errors:[
        "#VALUE!",
        "Invalid range"
    ],

    alternative:"=COUNTA(A1:F1)"
},

// ==========================================================
// FILTER
// ==========================================================
{
    title:"FILTER",

    keywords:[
        "filter","dynamic filter","extract data","show matching data",
        "remove unwanted data","conditional filter","advanced filter",
        "dynamic array","extract records","employee filter",
        "customer filter","sales filter","department filter",
        "status filter","find records","display matching rows",
        "return multiple results","dynamic report","interactive report",
        "excel 365","excel 2021","modern excel","data extraction",
        "table filtering","criteria based filter","search records",
        "multiple criteria","dashboard data","report automation",
        "database filter","conditional extraction","filtered list",
        "data analysis","record selection","live report"
    ],

    aliases:[
        "dynamic filtering",
        "extract matching rows",
        "show selected records",
        "criteria filter",
        "advanced filtering",
        "automatic filtering",
        "dynamic report formula",
        "filtered data list",
        "smart filter",
        "excel filter formula"
    ],

    questions:[
        "How to filter data using formula?",
        "How to extract matching records?",
        "How to filter employee data?",
        "How to show only selected rows?",
        "How to create dynamic reports?",
        "How to filter based on condition?",
        "How to return multiple results?",
        "How to filter sales data?",
        "How to create automatic filtering?",
        "How to use FILTER function?"
    ],

    dataset:"employees",

    category:"Dynamic Array",

    formula:'=FILTER(A2:E100,C2:C100="Finance")',

    syntax:"FILTER(array,include,[if_empty])",

    explanation:"Returns only the rows or columns that meet specified conditions.",

    breakdown:[
        "A2:E100 = Data Range",
        "C2:C100 = Condition Range",
        "Finance = Filter Criteria"
    ],

    output:"Finance Department Employees",

    business:"Create dynamic employee reports, sales analysis, and automated dashboards.",

    tips:[
        "Available in Excel 365 and Excel 2021.",
        "Can use multiple conditions.",
        "Works with dynamic arrays."
    ],

    errors:[
        "#CALC!",
        "No matching records",
        "Wrong criteria range"
    ],

    alternative:'=INDEX(A:A,MATCH("Finance",C:C,0))'
},


// ==========================================================
// SORT
// ==========================================================
{
    title:"SORT",

    keywords:[
        "sort","automatic sorting","dynamic sort",
        "arrange data","ascending order","descending order",
        "rank data","order list","sort records",
        "employee ranking","sales ranking","price sorting",
        "date sorting","alphabetical order","numeric sorting",
        "dynamic array","excel 365","excel 2021",
        "modern excel","data organization",
        "sorted report","automatic report",
        "dashboard sorting","table sorting",
        "rearrange data","sort by column",
        "sort by value","data analysis",
        "ranking report","business report",
        "dynamic table","live sorting"
    ],

    aliases:[
        "automatic sort",
        "dynamic sorting",
        "arrange records",
        "order data",
        "ranking formula",
        "sort list",
        "smart sorting",
        "excel sort formula",
        "sorted report",
        "data arrangement"
    ],

    questions:[
        "How to sort data using formula?",
        "How to arrange data automatically?",
        "How to sort employee salary?",
        "How to rank sales data?",
        "How to create dynamic sorted list?",
        "How to sort dates?",
        "How to sort alphabetically?",
        "How to sort by column?",
        "How to create automatic ranking?",
        "How to use SORT function?"
    ],

    dataset:"sales",

    category:"Dynamic Array",

    formula:"=SORT(A2:E100,4,-1)",

    syntax:"SORT(array,[sort_index],[sort_order],[by_col])",

    explanation:"Sorts a range or array dynamically based on selected column and order.",

    breakdown:[
        "A2:E100 = Data Range",
        "4 = Salary/Sales Column",
        "-1 = Descending Order"
    ],

    output:"Highest Sales First",

    business:"Create automatically sorted reports, rankings, and dashboards.",

    tips:[
        "Works without manual sorting.",
        "Supports ascending and descending order.",
        "Updates automatically when data changes."
    ],

    errors:[
        "#VALUE!",
        "Invalid sort index",
        "Wrong range"
    ],

    alternative:"=SORTBY(A2:E100,D2:D100,-1)"
},

// ==========================================================
// SORTBY
// ==========================================================
{
    title:"SORTBY",

    keywords:[
        "sortby","advanced sort","sort based on another column",
        "dynamic sorting","custom sorting","sort records",
        "sort sales","sort salary","sort date",
        "sort employee data","sort customer data",
        "multiple sorting","dynamic array","excel 365",
        "excel 2021","modern excel","automatic sorting",
        "ranking data","arrange table","order data",
        "sort by value","sort by criteria","business report",
        "dashboard sorting","data analysis","report automation",
        "ascending sorting","descending sorting",
        "multi column sorting","smart sorting",
        "dynamic report","live sorting","table sorting"
    ],

    aliases:[
        "advanced sorting formula",
        "sort using another column",
        "multi level sorting",
        "dynamic order formula",
        "custom sort formula",
        "smart sort",
        "automatic ranking",
        "sort table dynamically",
        "business sorting",
        "excel sorting function"
    ],

    questions:[
        "How to sort data based on another column?",
        "How to sort sales from highest to lowest?",
        "How to sort employee salary?",
        "How to create dynamic sorting?",
        "How to sort multiple columns?",
        "How to rank records automatically?",
        "How to sort by date?",
        "How to arrange data using formula?",
        "How to use SORTBY function?",
        "How to create sorted dashboard?"
    ],

    dataset:"sales",

    category:"Dynamic Array",

    formula:"=SORTBY(A2:E100,D2:D100,-1)",

    syntax:"SORTBY(array,by_array,[sort_order])",

    explanation:"Sorts a range based on values from another range or column.",

    breakdown:[
        "A2:E100 = Data Range",
        "D2:D100 = Sorting Column",
        "-1 = Descending Order"
    ],

    output:"Highest Sales Records First",

    business:"Create dynamic ranking reports, sales dashboards, and automated MIS reports.",

    tips:[
        "Can sort using multiple columns.",
        "Updates automatically when data changes.",
        "Better than manual sorting."
    ],

    errors:[
        "#VALUE!",
        "Range size mismatch",
        "Invalid sort order"
    ],

    alternative:"=SORT(A2:E100,4,-1)"
},


// ==========================================================
// UNIQUE
// ==========================================================
{
    title:"UNIQUE",

    keywords:[
        "unique","remove duplicate","duplicate removal",
        "distinct values","unique list","duplicate finder",
        "deduplicate data","clean data","data cleaning",
        "employee names","customer list","product list",
        "sales category","department list","master list",
        "dynamic array","excel 365","excel 2021",
        "modern excel","extract unique records",
        "distinct records","data preparation",
        "database cleanup","duplicate check",
        "list creation","automatic list",
        "report preparation","dashboard filter",
        "unique values","data quality",
        "remove repeated values","clean dataset"
    ],

    aliases:[
        "duplicate remover",
        "distinct list generator",
        "clean data formula",
        "unique value extractor",
        "deduplicate formula",
        "duplicate cleanup",
        "master list creator",
        "distinct record finder",
        "data cleaning formula",
        "excel unique function"
    ],

    questions:[
        "How to remove duplicates using formula?",
        "How to get unique values?",
        "How to create employee list?",
        "How to extract distinct records?",
        "How to clean duplicate data?",
        "How to create unique customer list?",
        "How to find unique products?",
        "How to generate dropdown list?",
        "How to use UNIQUE function?",
        "How to prepare clean data?"
    ],

    dataset:"customers",

    category:"Dynamic Array",

    formula:"=UNIQUE(A2:A100)",

    syntax:"UNIQUE(array,[by_col],[exactly_once])",

    explanation:"Returns only unique values from a range and removes duplicates dynamically.",

    breakdown:[
        "A2:A100 = Source Data",
        "Returns distinct values only"
    ],

    output:"Unique Customer Names",

    business:"Clean customer databases, create dropdown lists, and prepare analytical datasets.",

    tips:[
        "Works automatically when data changes.",
        "Can return values appearing once only.",
        "Useful before creating reports."
    ],

    errors:[
        "#CALC!",
        "Empty source range",
        "Incorrect array"
    ],

    alternative:"=Remove Duplicates feature"
},


// ==========================================================
// TAKE
// ==========================================================
{
    title:"TAKE",

    keywords:[
        "take","top rows","bottom rows","extract rows",
        "first records","latest records","sample data",
        "dynamic array","excel 365","modern excel",
        "return first values","return last values",
        "top sales","highest records","recent data",
        "limit rows","extract data","partial dataset",
        "report summary","dashboard summary",
        "quick analysis","data preview",
        "table extraction","dynamic report",
        "select rows","select columns",
        "first entries","last entries",
        "business analysis","data slicing",
        "record extraction","data selection"
    ],

    aliases:[
        "top records formula",
        "extract first rows",
        "extract last rows",
        "data slicer",
        "row extractor",
        "quick data preview",
        "dynamic sample",
        "record selector",
        "limited dataset",
        "excel take function"
    ],

    questions:[
        "How to get top 10 records?",
        "How to extract first rows?",
        "How to get latest records?",
        "How to select limited data?",
        "How to create summary report?",
        "How to extract bottom rows?",
        "How to preview dataset?",
        "How to use TAKE function?",
        "How to get sample records?",
        "How to create dynamic summary?"
    ],

    dataset:"sales",

    category:"Dynamic Array",

    formula:"=TAKE(A2:E100,10)",

    syntax:"TAKE(array,rows,[columns])",

    explanation:"Returns a specified number of rows or columns from the beginning or end of an array.",

    breakdown:[
        "A2:E100 = Source Data",
        "10 = Number of Rows Returned"
    ],

    output:"Top 10 Sales Records",

    business:"Create quick summaries, dashboards, and top performer reports.",

    tips:[
        "Use negative number for bottom records.",
        "Works with dynamic arrays.",
        "Useful for dashboard summaries."
    ],

    errors:[
        "#CALC!",
        "Requested rows exceed data",
        "Invalid array"
    ],

    alternative:"=INDEX(A2:E100,SEQUENCE(10),SEQUENCE(,5))"
},

// ==========================================================
// DROP
// ==========================================================
{
    title:"DROP",

    keywords:[
        "drop","remove rows","remove columns","exclude data",
        "remove header","skip records","remove first rows",
        "remove last rows","dynamic array","excel 365",
        "excel 2021","modern excel","data cleaning",
        "extract remaining data","remove unwanted records",
        "table cleanup","dataset preparation",
        "remove top rows","remove bottom rows",
        "remove columns","dynamic filtering",
        "report preparation","dashboard data",
        "data transformation","data manipulation",
        "clean dataset","remove blank rows",
        "prepare analysis data","excel automation",
        "dynamic report","range transformation"
    ],

    aliases:[
        "remove records formula",
        "exclude rows",
        "skip rows",
        "data cleaner",
        "remove columns dynamically",
        "dynamic remove function",
        "table trimming",
        "dataset cleaner",
        "remove unwanted data",
        "excel drop function"
    ],

    questions:[
        "How to remove first rows using formula?",
        "How to remove header dynamically?",
        "How to exclude unwanted columns?",
        "How to remove last records?",
        "How to clean data automatically?",
        "How to prepare data for analysis?",
        "How to remove rows from table?",
        "How to use DROP function?",
        "How to create cleaned dataset?",
        "How to remove top records?"
    ],

    dataset:"sales",

    category:"Dynamic Array",

    formula:"=DROP(A1:E100,1)",

    syntax:"DROP(array,rows,[columns])",

    explanation:"Removes a specified number of rows or columns from the beginning or end of an array.",

    breakdown:[
        "A1:E100 = Source Data",
        "1 = Remove First Row",
        "Returns remaining dataset"
    ],

    output:"Data Without Header",

    business:"Clean imported data and prepare datasets before analysis.",

    tips:[
        "Use negative numbers to remove from bottom/right.",
        "Useful with FILTER and SORT.",
        "Available in Excel 365."
    ],

    errors:[
        "#CALC!",
        "Removing more rows than available",
        "Invalid range"
    ],

    alternative:"=FILTER(A2:E100,A2:A100<>\"\")"
},


// ==========================================================
// CHOOSEROWS
// ==========================================================
{
    title:"CHOOSEROWS",

    keywords:[
        "chooserows","choose rows","select rows",
        "extract rows","pick rows","specific rows",
        "dynamic row selection","row extraction",
        "custom report","dashboard rows",
        "employee records","sales records",
        "specific data","selected records",
        "dynamic array","excel 365","modern excel",
        "table selection","data extraction",
        "report creation","business report",
        "row filtering","record selection",
        "analysis data","custom dataset",
        "extract selected rows","row management",
        "data slicing","advanced excel",
        "dynamic reporting","MIS report"
    ],

    aliases:[
        "row selector",
        "pick specific rows",
        "custom row extraction",
        "selected records",
        "dynamic row picker",
        "report row selector",
        "data slice formula",
        "advanced filtering",
        "row extraction formula",
        "excel row chooser"
    ],

    questions:[
        "How to select specific rows?",
        "How to extract selected records?",
        "How to create custom reports?",
        "How to pick rows from table?",
        "How to select employee records?",
        "How to create dashboard data?",
        "How to use CHOOSEROWS?",
        "How to extract required rows only?",
        "How to create dynamic reports?",
        "How to choose multiple rows?"
    ],

    dataset:"employees",

    category:"Dynamic Array",

    formula:"=CHOOSEROWS(A2:E100,1,5,10)",

    syntax:"CHOOSEROWS(array,row_num1,[row_num2])",

    explanation:"Returns selected rows from an array based on row numbers.",

    breakdown:[
        "A2:E100 = Source Data",
        "1,5,10 = Required Row Numbers",
        "Returns only selected rows"
    ],

    output:"Selected Employee Records",

    business:"Create customized reports by selecting only required records.",

    tips:[
        "Can select multiple rows.",
        "Works with dynamic arrays.",
        "Useful for dashboard summaries."
    ],

    errors:[
        "#VALUE!",
        "Invalid row number",
        "Row does not exist"
    ],

    alternative:"=FILTER(A2:E100,ROW(A2:A100)<=10)"
},


// ==========================================================
// CHOOSECOLS
// ==========================================================
{
    title:"CHOOSECOLS",

    keywords:[
        "choosecols","choose columns","select columns",
        "extract columns","pick columns","specific fields",
        "remove unwanted columns","custom report",
        "employee name salary","data extraction",
        "column selection","dynamic columns",
        "dashboard fields","table transformation",
        "dynamic array","excel 365","modern excel",
        "report preparation","business analysis",
        "data cleaning","database fields",
        "select required columns","MIS report",
        "custom dataset","column filtering",
        "data modelling","report automation",
        "advanced excel","field selection",
        "table formatting"
    ],

    aliases:[
        "column selector",
        "pick required fields",
        "custom column extraction",
        "field selector",
        "dynamic column picker",
        "report column selection",
        "data transformation formula",
        "remove unwanted fields",
        "excel column chooser",
        "column extraction"
    ],

    questions:[
        "How to select specific columns?",
        "How to extract required fields?",
        "How to create custom reports?",
        "How to remove unwanted columns?",
        "How to select employee name and salary?",
        "How to prepare dashboard data?",
        "How to use CHOOSECOLS?",
        "How to extract columns dynamically?",
        "How to create simplified tables?",
        "How to choose multiple columns?"
    ],

    dataset:"employees",

    category:"Dynamic Array",

    formula:"=CHOOSECOLS(A1:E100,1,4)",

    syntax:"CHOOSECOLS(array,col_num1,[col_num2])",

    explanation:"Returns selected columns from an array based on column numbers.",

    breakdown:[
        "A1:E100 = Source Data",
        "1 = Employee Name Column",
        "4 = Salary Column"
    ],

    output:"Employee Name and Salary Table",

    business:"Create clean dashboards and reports with only required fields.",

    tips:[
        "Useful before creating charts.",
        "Reduces unnecessary columns.",
        "Works well with FILTER."
    ],

    errors:[
        "#VALUE!",
        "Invalid column number",
        "Column does not exist"
    ],

    alternative:"=FILTER(A:E,A:A<>\"\")"
},

// ==========================================================
// WRAPROWS
// ==========================================================
{
    title:"WRAPROWS",

    keywords:[
        "wraprows","wrap rows","convert column to rows",
        "horizontal array","reshape data","format array",
        "dynamic array","excel 365","modern excel",
        "convert list","create rows","arrange values",
        "transform data","data layout","table formatting",
        "array transformation","data reshaping",
        "horizontal display","list to table",
        "dashboard formatting","report formatting",
        "organize data","matrix creation",
        "row arrangement","excel automation",
        "dynamic report","data presentation",
        "convert vertical data","array layout",
        "structured output","advanced excel"
    ],

    aliases:[
        "column to row converter",
        "array wrapper",
        "horizontal data formatter",
        "list formatter",
        "data reshaper",
        "row creator",
        "array organizer",
        "excel layout function",
        "dynamic row generator",
        "data arrangement formula"
    ],

    questions:[
        "How to convert column into rows?",
        "How to arrange values horizontally?",
        "How to create rows from a list?",
        "How to reshape Excel data?",
        "How to format array output?",
        "How to use WRAPROWS function?",
        "How to convert vertical data?",
        "How to create matrix layout?",
        "How to organize dynamic arrays?",
        "How to display data in rows?"
    ],

    dataset:"sales",

    category:"Dynamic Array",

    formula:"=WRAPROWS(A2:A13,3)",

    syntax:"WRAPROWS(vector,wrap_count,[pad_with])",

    explanation:"Wraps values from a single row or column into multiple rows.",

    breakdown:[
        "A2:A13 = Source List",
        "3 = Number of Values Per Row",
        "Creates horizontal groups"
    ],

    output:"3 Values Per Row",

    business:"Convert long lists into formatted reports and dashboard layouts.",

    tips:[
        "Useful for presentation formatting.",
        "Works with dynamic arrays.",
        "Can add padding values."
    ],

    errors:[
        "#VALUE!",
        "Invalid wrap count"
    ],

    alternative:"=TRANSPOSE(A2:A13)"
},


// ==========================================================
// WRAPCOLS
// ==========================================================
{
    title:"WRAPCOLS",

    keywords:[
        "wrapcols","wrap columns","convert row to columns",
        "vertical array","reshape data","column format",
        "dynamic array","excel 365","modern excel",
        "convert horizontal data","create columns",
        "array transformation","data formatting",
        "table layout","report formatting",
        "matrix creation","data organization",
        "vertical display","list transformation",
        "dashboard layout","structured data",
        "excel automation","dynamic report",
        "column arrangement","array management",
        "data presentation","advanced excel",
        "convert rows","reshape table",
        "column generator","format values"
    ],

    aliases:[
        "row to column converter",
        "vertical data formatter",
        "column creator",
        "array organizer",
        "data reshaping function",
        "dynamic column generator",
        "table formatting formula",
        "excel layout function",
        "column wrapper",
        "data transformation formula"
    ],

    questions:[
        "How to convert rows into columns?",
        "How to arrange data vertically?",
        "How to create columns from a list?",
        "How to reshape Excel data?",
        "How to use WRAPCOLS?",
        "How to format horizontal data?",
        "How to create vertical arrays?",
        "How to organize reports?",
        "How to transform data layout?",
        "How to convert row values?"
    ],

    dataset:"products",

    category:"Dynamic Array",

    formula:"=WRAPCOLS(A2:L2,4)",

    syntax:"WRAPCOLS(vector,wrap_count,[pad_with])",

    explanation:"Wraps values into columns based on the specified number of values.",

    breakdown:[
        "A2:L2 = Source Row",
        "4 = Values Per Column",
        "Creates vertical groups"
    ],

    output:"Product List By Column",

    business:"Format horizontal data into structured vertical reports.",

    tips:[
        "Useful with dashboard layouts.",
        "Works with dynamic arrays."
    ],

    errors:[
        "#VALUE!",
        "Invalid wrap count"
    ],

    alternative:"=TRANSPOSE(A2:L2)"
},


// ==========================================================
// TOCOL
// ==========================================================
{
    title:"TOCOL",

    keywords:[
        "tocol","convert to column","flatten data",
        "combine ranges","single column","stack data",
        "merge lists","dynamic array","excel 365",
        "remove empty cells","data transformation",
        "reshape data","normalize data","clean data",
        "table conversion","vertical list",
        "combine tables","data preparation",
        "power query alternative","array function",
        "multiple range merge","report preparation",
        "data analysis","excel automation",
        "list creation","database preparation",
        "flatten array","convert matrix",
        "structured data","advanced excel"
    ],

    aliases:[
        "flatten array",
        "merge lists",
        "combine columns",
        "single column generator",
        "data stacker",
        "range combiner",
        "list converter",
        "array cleaner",
        "vertical converter",
        "data normalization"
    ],

    questions:[
        "How to combine multiple columns into one?",
        "How to flatten Excel data?",
        "How to create a single list?",
        "How to merge ranges?",
        "How to convert matrix to column?",
        "How to clean multiple ranges?",
        "How to use TOCOL function?",
        "How to prepare data for analysis?",
        "How to stack data automatically?",
        "How to create one column list?"
    ],

    dataset:"sales",

    category:"Dynamic Array",

    formula:"=TOCOL(A2:C10)",

    syntax:"TOCOL(array,[ignore],[scan_by_column])",

    explanation:"Converts an array or range into a single column.",

    breakdown:[
        "A2:C10 = Source Range",
        "Returns all values in one column"
    ],

    output:"Combined Data List",

    business:"Prepare data for analysis by converting multiple columns into one list.",

    tips:[
        "Useful before UNIQUE and FILTER.",
        "Helps normalize data."
    ],

    errors:[
        "#VALUE!",
        "Invalid array"
    ],

    alternative:"=VSTACK(A:A,B:B,C:C)"
},

// ==========================================================
// TOROW
// ==========================================================
{
    title:"TOROW",

    keywords:[
        "torow","convert to row","flatten data",
        "combine ranges","single row","merge lists",
        "horizontal list","dynamic array","excel 365",
        "data transformation","reshape data",
        "array conversion","convert column to row",
        "table transformation","data preparation",
        "clean data","report formatting",
        "dashboard layout","multiple range merge",
        "horizontal output","list generator",
        "data analysis","excel automation",
        "matrix conversion","array function",
        "structured data","data normalization",
        "combine values","range conversion",
        "advanced excel","modern excel"
    ],

    aliases:[
        "flatten to row",
        "merge columns into row",
        "horizontal converter",
        "single row generator",
        "array flattener",
        "list converter",
        "range combiner",
        "data reshape formula",
        "horizontal data list",
        "excel conversion function"
    ],

    questions:[
        "How to combine data into one row?",
        "How to convert columns into rows?",
        "How to flatten Excel data horizontally?",
        "How to merge multiple ranges?",
        "How to create a single row list?",
        "How to use TOROW function?",
        "How to transform array data?",
        "How to prepare dashboard data?",
        "How to convert matrix into row?",
        "How to create horizontal output?"
    ],

    dataset:"sales",

    category:"Dynamic Array",

    formula:"=TOROW(A2:C10)",

    syntax:"TOROW(array,[ignore],[scan_by_column])",

    explanation:"Converts an array or range into a single row.",

    breakdown:[
        "A2:C10 = Source Range",
        "Returns all values in one row"
    ],

    output:"Combined Horizontal Data",

    business:"Create dashboard layouts and transform tables into reporting formats.",

    tips:[
        "Useful with UNIQUE and SORT.",
        "Helps reshape imported data.",
        "Works only in modern Excel."
    ],

    errors:[
        "#VALUE!",
        "Invalid array",
        "Unsupported Excel version"
    ],

    alternative:"=TRANSPOSE(A2:A100)"
},


// ==========================================================
// EXPAND
// ==========================================================
{
    title:"EXPAND",

    keywords:[
        "expand","increase array","add rows",
        "add columns","resize array","dynamic array",
        "excel 365","modern excel","extend table",
        "add blank values","increase range",
        "array resize","table expansion",
        "dynamic report","dashboard formatting",
        "matrix expansion","data formatting",
        "fill missing values","pad array",
        "extend dataset","report preparation",
        "structured output","custom table size",
        "dynamic layout","array manipulation",
        "excel automation","data transformation",
        "resize data","add placeholders",
        "advanced excel","formula expansion"
    ],

    aliases:[
        "array extender",
        "resize table",
        "increase range size",
        "dynamic array resize",
        "add empty cells",
        "expand dataset",
        "table enlarger",
        "array padding",
        "extend formula output",
        "dynamic table creator"
    ],

    questions:[
        "How to expand array size?",
        "How to add rows dynamically?",
        "How to increase columns using formula?",
        "How to create fixed size output?",
        "How to add blank values?",
        "How to resize dynamic arrays?",
        "How to use EXPAND function?",
        "How to format dashboard tables?",
        "How to create consistent report size?",
        "How to extend formula results?"
    ],

    dataset:"dashboard",

    category:"Dynamic Array",

    formula:'=EXPAND(A2:B5,10,5,"")',

    syntax:"EXPAND(array,rows,[columns],[pad_with])",

    explanation:"Expands an array to a specified number of rows and columns with optional padding values.",

    breakdown:[
        "A2:B5 = Original Array",
        "10 = Required Rows",
        "5 = Required Columns",
        "\"\" = Empty Padding"
    ],

    output:"Expanded Report Table",

    business:"Create fixed-size dashboards and standardized reports.",

    tips:[
        "Useful when creating dashboard templates.",
        "Can add custom padding values.",
        "Works with other dynamic array functions."
    ],

    errors:[
        "#VALUE!",
        "Invalid row size",
        "Invalid column size"
    ],

    alternative:"=VSTACK(A2:B5,extra_range)"
},

// ==========================================================
// IF
// ==========================================================
{
    title:"IF",

    keywords:[
        "if","if formula","condition","conditional formula",
        "logical test","check condition","true false",
        "yes no","pass fail","status check",
        "employee status","salary condition","bonus calculation",
        "approval","rejection","validation","decision formula",
        "compare values","greater than","less than",
        "equal condition","text condition","number condition",
        "business rule","automation","data validation",
        "grading","performance rating","attendance",
        "leave status","target achievement","commission",
        "excel logic","logical function","decision making"
    ],

    aliases:[
        "conditional formula",
        "decision formula",
        "true false formula",
        "yes no formula",
        "status formula",
        "check value formula",
        "business rule formula",
        "condition checker",
        "logical test formula",
        "excel decision function"
    ],

    questions:[
        "How to use IF formula?",
        "How to create pass fail formula?",
        "How to check employee status?",
        "How to calculate bonus based on condition?",
        "How to apply condition in Excel?",
        "How to return yes or no?",
        "How to compare two values?",
        "How to create approval logic?",
        "How to check attendance status?",
        "How to create conditional calculation?"
    ],

    dataset:"employees",

    category:"Logical",

    formula:'=IF(C2>=50000,"High Salary","Low Salary")',

    syntax:"IF(logical_test,value_if_true,value_if_false)",

    explanation:"Checks a condition and returns one result when TRUE and another result when FALSE.",

    breakdown:[
        "C2>=50000 = Condition",
        "\"High Salary\" = TRUE Result",
        "\"Low Salary\" = FALSE Result"
    ],

    output:"High Salary",

    business:"Used for employee grading, approvals, status tracking, bonus calculation, and business decisions.",

    tips:[
        "Combine IF with AND/OR for multiple conditions.",
        "Use IFERROR to handle formula errors.",
        "Avoid too many nested IF formulas."
    ],

    errors:[
        "#VALUE!",
        "Wrong logical condition",
        "Missing quotation marks"
    ],

    alternative:'=IFS(C2>=50000,"High Salary",C2<50000,"Low Salary")'
},


// ==========================================================
// IFS
// ==========================================================
{
    title:"IFS",

    keywords:[
        "ifs","multiple if","multiple conditions",
        "nested if alternative","condition checking",
        "multiple criteria","grading formula",
        "employee rating","performance rating",
        "salary band","commission calculation",
        "bonus slab","approval levels",
        "category formula","classification",
        "score calculation","grade calculation",
        "business logic","decision tree",
        "logical test","multiple rules",
        "status mapping","rank calculation",
        "performance category","excel 365",
        "modern excel","replace nested if",
        "advanced condition","automation",
        "dynamic decision","rule based formula"
    ],

    aliases:[
        "multiple condition formula",
        "nested IF replacement",
        "grading formula",
        "rating formula",
        "rule based formula",
        "category generator",
        "decision tree formula",
        "multiple criteria checker",
        "business rule formula",
        "advanced IF formula"
    ],

    questions:[
        "How to check multiple conditions?",
        "How to replace nested IF?",
        "How to create grade formula?",
        "How to calculate performance rating?",
        "How to create salary bands?",
        "How to apply multiple rules?",
        "How to classify employees?",
        "How to create category formula?",
        "How to use IFS function?",
        "How to simplify multiple IF formulas?"
    ],

    dataset:"performance",

    category:"Logical",

    formula:'=IFS(B2>=90,"Excellent",B2>=75,"Good",B2>=50,"Average",TRUE,"Poor")',

    syntax:"IFS(test1,result1,test2,result2,...)",

    explanation:"Checks multiple conditions and returns the result for the first TRUE condition.",

    breakdown:[
        "B2>=90 = Excellent",
        "B2>=75 = Good",
        "B2>=50 = Average",
        "TRUE = Default Result"
    ],

    output:"Excellent",

    business:"Used for employee ratings, salary slabs, risk categories, and performance classification.",

    tips:[
        "Cleaner alternative to nested IF.",
        "Add TRUE at the end as a default condition.",
        "Available in Excel 2019 and Microsoft 365."
    ],

    errors:[
        "#VALUE!",
        "No TRUE condition found",
        "Incorrect condition order"
    ],

    alternative:'=IF(B2>=90,"Excellent",IF(B2>=75,"Good","Average"))'
},

// ==========================================================
// AND
// ==========================================================
{
    title:"AND",

    keywords:[
        "and","multiple conditions","all conditions",
        "check multiple criteria","logical test",
        "both conditions","all true","eligibility check",
        "employee eligibility","approval formula",
        "loan approval","salary criteria","attendance criteria",
        "performance check","validation formula",
        "combined condition","multiple rules",
        "business logic","decision formula",
        "true false","conditional check",
        "data validation","quality check",
        "compliance check","risk assessment",
        "excel logic","logical function",
        "pass criteria","requirement check",
        "multiple requirements","approval workflow"
    ],

    aliases:[
        "multiple criteria checker",
        "all conditions formula",
        "combined condition formula",
        "eligibility formula",
        "approval logic",
        "validation formula",
        "business rule checker",
        "logical AND formula",
        "criteria validator",
        "requirement checker"
    ],

    questions:[
        "How to check multiple conditions?",
        "How to verify all conditions are true?",
        "How to create eligibility formula?",
        "How to check employee requirements?",
        "How to combine conditions?",
        "How to create approval logic?",
        "How to use AND function?",
        "How to validate multiple criteria?",
        "How to check salary and performance conditions?",
        "How to create multiple rule checks?"
    ],

    dataset:"employees",

    category:"Logical",

    formula:'=AND(B2="Active",C2>=50000)',

    syntax:"AND(logical1,logical2,...)",

    explanation:"Returns TRUE only when all specified conditions are TRUE.",

    breakdown:[
        "B2=\"Active\" = Employee Status Check",
        "C2>=50000 = Salary Condition",
        "Both conditions must be TRUE"
    ],

    output:"TRUE",

    business:"Used for employee eligibility, approval workflows, compliance checks, and validation rules.",

    tips:[
        "Combine with IF for readable results.",
        "All conditions must pass.",
        "Useful for audit and control checks."
    ],

    errors:[
        "#VALUE!",
        "Wrong logical comparison",
        "Incorrect text condition"
    ],

    alternative:'=IF(AND(B2="Active",C2>=50000),"Eligible","Not Eligible")'
},


// ==========================================================
// OR
// ==========================================================
{
    title:"OR",

    keywords:[
        "or","any condition","one condition",
        "multiple options","alternative condition",
        "either condition","logical test",
        "employee status","approval rule",
        "exception handling","validation",
        "multiple criteria","business rule",
        "risk check","alert formula",
        "attendance status","leave approval",
        "performance check","condition formula",
        "true false","decision formula",
        "excel logic","logical function",
        "one match enough","criteria check",
        "multiple scenarios","workflow automation",
        "rule based formula","exception formula",
        "data quality check","control testing"
    ],

    aliases:[
        "any condition formula",
        "alternative criteria",
        "either option formula",
        "multiple option checker",
        "exception checker",
        "logical OR formula",
        "condition alternative",
        "rule validation",
        "scenario checker",
        "business exception formula"
    ],

    questions:[
        "How to check if any condition is true?",
        "How to create OR formula?",
        "How to check multiple options?",
        "How to validate alternative conditions?",
        "How to create exception rules?",
        "How to check employee status?",
        "How to use OR function?",
        "How to apply multiple choices?",
        "How to create alert formula?",
        "How to test different scenarios?"
    ],

    dataset:"employees",

    category:"Logical",

    formula:'=OR(B2="Manager",B2="Senior")',

    syntax:"OR(logical1,logical2,...)",

    explanation:"Returns TRUE when at least one condition is TRUE.",

    breakdown:[
        "B2=\"Manager\" = First Condition",
        "B2=\"Senior\" = Second Condition",
        "Any TRUE condition returns TRUE"
    ],

    output:"TRUE",

    business:"Used for exception checks, approval rules, alerts, and alternative scenarios.",

    tips:[
        "Combine with IF for output messages.",
        "Only one condition needs to be TRUE.",
        "Useful for exception reporting."
    ],

    errors:[
        "#VALUE!",
        "Invalid logical expression"
    ],

    alternative:'=IF(OR(B2="Manager",B2="Senior"),"Eligible","No")'
},


// ==========================================================
// NOT
// ==========================================================
{
    title:"NOT",

    keywords:[
        "not","reverse condition","opposite result",
        "negate condition","logical opposite",
        "exclude condition","not equal",
        "check exception","validation formula",
        "negative condition","status check",
        "error checking","compliance check",
        "data validation","filter exclusion",
        "logical function","true false",
        "business rule","exception logic",
        "remove condition","avoid match",
        "excel logic","decision formula",
        "reverse boolean","control testing",
        "quality check","audit check",
        "exception report","rule validation"
    ],

    aliases:[
        "reverse logic formula",
        "opposite condition",
        "exclude criteria",
        "negative condition",
        "exception checker",
        "reverse boolean",
        "logical inverter",
        "not equal checker",
        "condition reverser",
        "validation opposite"
    ],

    questions:[
        "How to reverse a condition?",
        "How to check not equal?",
        "How to exclude values?",
        "How to create opposite logic?",
        "How to use NOT function?",
        "How to find exceptions?",
        "How to validate negative conditions?",
        "How to ignore specific values?",
        "How to create exclusion formula?",
        "How to reverse TRUE FALSE?"
    ],

    dataset:"employees",

    category:"Logical",

    formula:'=NOT(B2="Inactive")',

    syntax:"NOT(logical)",

    explanation:"Reverses the logical value. TRUE becomes FALSE and FALSE becomes TRUE.",

    breakdown:[
        "B2=\"Inactive\" = Original Condition",
        "NOT reverses the result"
    ],

    output:"TRUE",

    business:"Used for exception reports, validation checks, and exclusion rules.",

    tips:[
        "Useful with IF, AND, and OR.",
        "Good for checking exceptions.",
        "Can simplify negative logic."
    ],

    errors:[
        "#VALUE!",
        "Invalid logical value"
    ],

    alternative:'=B2<>"Inactive"'
},


// ==========================================================
// XOR
// ==========================================================
{
    title:"XOR",

    keywords:[
        "xor","exclusive or","only one condition",
        "one true condition","logical comparison",
        "either but not both","exclusive logic",
        "validation rule","decision formula",
        "multiple conditions","logical function",
        "boolean logic","condition testing",
        "business rule","exception handling",
        "approval logic","data validation",
        "scenario analysis","control testing",
        "audit logic","rule engine",
        "advanced excel","true false",
        "logical formula","workflow logic",
        "one option selected","exclusive selection",
        "condition checker","excel logic"
    ],

    aliases:[
        "exclusive OR formula",
        "only one condition",
        "single choice logic",
        "exclusive condition",
        "boolean checker",
        "advanced logical formula",
        "one option validator",
        "selection rule",
        "exclusive validation",
        "logical comparison formula"
    ],

    questions:[
        "How to check only one condition?",
        "How to use XOR formula?",
        "How to allow one option only?",
        "How to compare two conditions?",
        "How to create exclusive logic?",
        "How to check one true result?",
        "How to validate selections?",
        "How to create advanced conditions?",
        "How to use XOR in Excel?",
        "How to test exclusive scenarios?"
    ],

    dataset:"validation",

    category:"Logical",

    formula:"=XOR(A2=\"Yes\",B2=\"Yes\")",

    syntax:"XOR(logical1,logical2,...)",

    explanation:"Returns TRUE when an odd number of conditions are TRUE.",

    breakdown:[
        "Condition 1 = A2 value check",
        "Condition 2 = B2 value check",
        "Only one TRUE returns TRUE"
    ],

    output:"TRUE",

    business:"Useful for exclusive selections, validation rules, and workflow controls.",

    tips:[
        "Useful when only one option should be selected.",
        "Works with multiple logical tests.",
        "Available in modern Excel versions."
    ],

    errors:[
        "#VALUE!",
        "Invalid logical condition"
    ],

    alternative:"=OR(A2<>B2)"
},

// ==========================================================
// IFERROR
// ==========================================================
{
    title:"IFERROR",

    keywords:[
        "iferror","error handling","remove error",
        "handle error","replace error message",
        "#n/a fix","#value fix","#div0 fix",
        "avoid error","error formula",
        "display message instead error",
        "lookup error","vlookup error",
        "xlookup error","calculation error",
        "formula validation","clean report",
        "professional excel","dashboard formula",
        "data quality","error management",
        "excel automation","report cleanup",
        "prevent errors","blank instead error",
        "default value","exception handling",
        "formula protection","business report",
        "error checking","advanced excel"
    ],

    aliases:[
        "error handler",
        "error remover",
        "formula protection",
        "safe formula",
        "error replacement",
        "clean formula",
        "error management formula",
        "exception handler",
        "fallback formula",
        "error control"
    ],

    questions:[
        "How to remove Excel errors?",
        "How to replace #N/A with text?",
        "How to handle formula errors?",
        "How to show blank instead of error?",
        "How to fix VLOOKUP error?",
        "How to create error free reports?",
        "How to use IFERROR function?",
        "How to handle calculation errors?",
        "How to display custom messages?",
        "How to protect formulas?"
    ],

    dataset:"reports",

    category:"Logical",

    formula:'=IFERROR(VLOOKUP(H2,A:E,5,FALSE),"Not Found")',

    syntax:"IFERROR(value,value_if_error)",

    explanation:"Returns a custom result when a formula produces an error.",

    breakdown:[
        "VLOOKUP = Formula to test",
        "\"Not Found\" = Error Result",
        "Returns normal result if no error"
    ],

    output:"Not Found",

    business:"Create clean MIS reports and prevent visible formula errors.",

    tips:[
        "Use with lookup formulas.",
        "Improves dashboard presentation.",
        "Avoid hiding important errors during audits."
    ],

    errors:[
        "Incorrect error handling",
        "Hidden calculation issues"
    ],

    alternative:'=IFNA(VLOOKUP(H2,A:E,5,FALSE),"Not Found")'
},


// ==========================================================
// IFNA
// ==========================================================
{
    title:"IFNA",

    keywords:[
        "ifna","na error","n/a handling",
        "#n/a fix","lookup error","missing value",
        "not found message","vlookup not found",
        "xlookup missing result","handle missing data",
        "replace na","data validation",
        "clean lookup","professional report",
        "error management","missing record",
        "database lookup","excel formula",
        "report automation","data quality",
        "exception handling","lookup protection",
        "master data","record missing",
        "formula cleanup","dashboard error",
        "excel automation","advanced lookup",
        "error replacement"
    ],

    aliases:[
        "NA error handler",
        "missing value handler",
        "lookup error fix",
        "not found handler",
        "missing record formula",
        "lookup protection",
        "clean lookup formula",
        "N/A remover",
        "record validation",
        "search error handler"
    ],

    questions:[
        "How to fix #N/A error?",
        "How to show message when value missing?",
        "How to handle VLOOKUP not found?",
        "How to replace missing data?",
        "How to use IFNA function?",
        "How to clean lookup reports?",
        "How to handle missing records?",
        "How to avoid N/A errors?",
        "How to create lookup protection?",
        "How to manage missing values?"
    ],

    dataset:"employees",

    category:"Logical",

    formula:'=IFNA(XLOOKUP(H2,A:A,D:D),"Employee Not Found")',

    syntax:"IFNA(value,value_if_na)",

    explanation:"Handles only #N/A errors and returns a custom result.",

    breakdown:[
        "XLOOKUP = Lookup Formula",
        "Employee Not Found = Missing Value Message"
    ],

    output:"Employee Not Found",

    business:"Useful for employee lookup, customer databases, and master data validation.",

    tips:[
        "Use when only #N/A needs handling.",
        "Keeps other formula errors visible.",
        "Better than IFERROR for audits."
    ],

    errors:[
        "Incorrect lookup range",
        "Missing lookup value"
    ],

    alternative:'=IFERROR(XLOOKUP(H2,A:A,D:D),"Employee Not Found")'
},


// ==========================================================
// SWITCH
// ==========================================================
{
    title:"SWITCH",

    keywords:[
        "switch","replace nested if","multiple choices",
        "value mapping","category mapping",
        "status mapping","employee grade",
        "department mapping","rating formula",
        "classification","condition mapping",
        "alternative if","decision formula",
        "business rules","workflow logic",
        "multiple results","text matching",
        "score mapping","commission slab",
        "bonus calculation","excel logic",
        "advanced formula","automation",
        "dynamic decision","reporting",
        "category formula","rule based formula",
        "modern excel","decision table"
    ],

    aliases:[
        "nested if replacement",
        "mapping formula",
        "value selector",
        "category converter",
        "decision mapper",
        "status converter",
        "classification formula",
        "rule mapping",
        "choice formula",
        "excel switch formula"
    ],

    questions:[
        "How to replace nested IF?",
        "How to map values to categories?",
        "How to create status formula?",
        "How to assign grades?",
        "How to use SWITCH function?",
        "How to simplify multiple conditions?",
        "How to create category mapping?",
        "How to create rating formula?",
        "How to map department names?",
        "How to create decision logic?"
    ],

    dataset:"employees",

    category:"Logical",

    formula:'=SWITCH(B2,"A","Excellent","B","Good","C","Average","Poor")',

    syntax:"SWITCH(expression,value1,result1,...,default)",

    explanation:"Compares one value against multiple options and returns the matching result.",

    breakdown:[
        "B2 = Value to check",
        "A/B/C = Matching options",
        "Returns corresponding result"
    ],

    output:"Excellent",

    business:"Used for grading, classification, status mapping, and business rules.",

    tips:[
        "Cleaner than multiple IF formulas.",
        "Add default value at the end.",
        "Best for exact matches."
    ],

    errors:[
        "#N/A",
        "Missing default value",
        "Incorrect mapping"
    ],

    alternative:'=IFS(B2="A","Excellent",B2="B","Good")'
},


// ==========================================================
// TRUE
// ==========================================================
{
    title:"TRUE",

    keywords:[
        "true","true value","logical true",
        "boolean value","condition result",
        "default condition","always true",
        "formula logic","if true",
        "logical testing","excel boolean",
        "decision formula","validation",
        "comparison result","rule checking",
        "default case","conditional logic",
        "advanced excel","formula testing",
        "status validation","business logic",
        "control formula","logical constant"
    ],

    aliases:[
        "boolean true",
        "logical constant",
        "always true condition",
        "default true value",
        "excel true function"
    ],

    questions:[
        "How to return TRUE value?",
        "How to create always true condition?",
        "How to use TRUE function?",
        "How to test logical results?",
        "How to create default condition?"
    ],

    dataset:"validation",

    category:"Logical",

    formula:"=TRUE()",

    syntax:"TRUE()",

    explanation:"Returns the logical value TRUE.",

    breakdown:[
        "No arguments required",
        "Returns TRUE value"
    ],

    output:"TRUE",

    business:"Used in logical tests, validation rules, and formulas requiring a TRUE condition.",

    tips:[
        "Often used as final condition in IFS.",
        "Can be written as TRUE without brackets."
    ],

    errors:[
        "No major errors"
    ],

    alternative:"=1=1"
},


// ==========================================================
// FALSE
// ==========================================================
{
    title:"FALSE",

    keywords:[
        "false","false value","logical false",
        "boolean false","negative result",
        "condition failure","default false",
        "formula logic","excel boolean",
        "validation","logical testing",
        "decision formula","comparison result",
        "rule checking","business logic",
        "control formula","advanced excel",
        "status validation","logical constant"
    ],

    aliases:[
        "boolean false",
        "logical constant",
        "default false value",
        "negative condition",
        "excel false function"
    ],

    questions:[
        "How to return FALSE value?",
        "How to create false condition?",
        "How to use FALSE function?",
        "How to test failed condition?",
        "How to create default false result?"
    ],

    dataset:"validation",

    category:"Logical",

    formula:"=FALSE()",

    syntax:"FALSE()",

    explanation:"Returns the logical value FALSE.",

    breakdown:[
        "No arguments required",
        "Returns FALSE value"
    ],

    output:"FALSE",

    business:"Used for logical testing, validation, and decision formulas.",

    tips:[
        "Useful for testing conditions.",
        "Can be used as a default result."
    ],

    errors:[
        "No major errors"
    ],

    alternative:"=1=0"
},

// ==========================================================
// LEFT
// ==========================================================
{
    title:"LEFT",

    keywords:[
        "left","left formula","extract text",
        "extract characters","get first characters",
        "first letters","first name extraction",
        "employee code","customer code",
        "product code","text extraction",
        "substring","text split","data cleaning",
        "remove unwanted text","prefix extraction",
        "starting characters","beginning text",
        "excel text function","string extraction",
        "data preparation","master data cleaning",
        "text manipulation","report formatting",
        "automation","excel automation",
        "identifier extraction","code extraction",
        "dynamic text","business data cleaning",
        "text analysis","advanced excel"
    ],

    aliases:[
        "first character extractor",
        "starting text extractor",
        "prefix remover",
        "left side text extraction",
        "beginning value extractor",
        "text splitter",
        "code extractor",
        "string extractor",
        "name extractor",
        "excel substring formula"
    ],

    questions:[
        "How to extract first characters?",
        "How to get first name from text?",
        "How to extract employee code?",
        "How to remove extra text?",
        "How to split text from left side?",
        "How to extract prefix?",
        "How to use LEFT function?",
        "How to clean text data?",
        "How to get first few letters?",
        "How to extract starting values?"
    ],

    dataset:"employees",

    category:"Text",

    formula:'=LEFT(A2,5)',

    syntax:"LEFT(text,[num_chars])",

    explanation:"Returns the specified number of characters from the beginning of a text string.",

    breakdown:[
        "A2 = Source Text",
        "5 = Number of Characters",
        "Returns first 5 characters"
    ],

    output:"EMP01",

    business:"Used to extract employee codes, prefixes, product IDs, and clean business data.",

    tips:[
        "Combine with LEN for dynamic extraction.",
        "Useful for cleaning imported data.",
        "Works with text and numbers stored as text."
    ],

    errors:[
        "#VALUE!",
        "Invalid character count",
        "Wrong cell reference"
    ],

    alternative:"=TEXTBEFORE(A2,\"-\")"
},


// ==========================================================
// RIGHT
// ==========================================================
{
    title:"RIGHT",

    keywords:[
        "right","right formula","extract ending text",
        "last characters","get last digits",
        "extract suffix","ending value",
        "account number","employee id",
        "phone number","invoice number",
        "text extraction","substring",
        "data cleaning","text split",
        "remove unwanted text","last letters",
        "excel text function","string extraction",
        "identifier extraction","code extraction",
        "report formatting","data preparation",
        "automation","excel automation",
        "suffix extraction","trailing characters",
        "text manipulation","business data",
        "advanced excel","text analysis"
    ],

    aliases:[
        "last character extractor",
        "ending text extractor",
        "suffix remover",
        "right side extraction",
        "last digit formula",
        "string end extractor",
        "code suffix extractor",
        "text splitter",
        "ending value finder",
        "excel substring formula"
    ],

    questions:[
        "How to extract last characters?",
        "How to get last digits?",
        "How to extract account number ending?",
        "How to split text from right side?",
        "How to extract suffix?",
        "How to get last few letters?",
        "How to use RIGHT function?",
        "How to clean ending text?",
        "How to extract invoice code?",
        "How to extract last values?"
    ],

    dataset:"transactions",

    category:"Text",

    formula:'=RIGHT(A2,4)',

    syntax:"RIGHT(text,[num_chars])",

    explanation:"Returns the specified number of characters from the end of a text string.",

    breakdown:[
        "A2 = Source Text",
        "4 = Number of Characters",
        "Returns last 4 characters"
    ],

    output:"4589",

    business:"Used for extracting account endings, invoice numbers, IDs, and transaction references.",

    tips:[
        "Useful for masked account numbers.",
        "Combine with LEN for dynamic formulas.",
        "Works well with MID."
    ],

    errors:[
        "#VALUE!",
        "Invalid character count",
        "Wrong reference"
    ],

    alternative:"=TEXTAFTER(A2,\"-\")"
},

// ==========================================================
// MID
// ==========================================================
{
    title:"MID",

    keywords:[
        "mid","mid formula","extract middle text",
        "extract characters from middle","substring",
        "text extraction","employee code",
        "product code","transaction id",
        "extract part of text","split text",
        "middle characters","data cleaning",
        "string extraction","text manipulation",
        "excel text function","code extraction",
        "identifier extraction","master data cleaning",
        "imported data cleaning","report preparation",
        "text parsing","business data",
        "automation","excel automation",
        "dynamic extraction","text analysis",
        "advanced excel","data preparation"
    ],

    aliases:[
        "middle text extractor",
        "substring extractor",
        "middle character formula",
        "text parser",
        "code separator",
        "string middle extractor",
        "partial text extractor",
        "data split formula",
        "text segment extractor",
        "excel substring function"
    ],

    questions:[
        "How to extract middle characters?",
        "How to get text from middle?",
        "How to split employee code?",
        "How to extract part of a string?",
        "How to use MID function?",
        "How to extract transaction number?",
        "How to separate text values?",
        "How to clean mixed data?",
        "How to extract characters between positions?",
        "How to parse text data?"
    ],

    dataset:"employees",

    category:"Text",

    formula:'=MID(A2,4,5)',

    syntax:"MID(text,start_num,num_chars)",

    explanation:"Returns a specific number of characters from the middle of a text string.",

    breakdown:[
        "A2 = Source Text",
        "4 = Starting Position",
        "5 = Number of Characters"
    ],

    output:"ABC12",

    business:"Used to extract codes, IDs, references, and specific sections from business data.",

    tips:[
        "Combine with FIND for dynamic extraction.",
        "Useful for fixed-format codes.",
        "Works with LEFT and RIGHT."
    ],

    errors:[
        "#VALUE!",
        "Invalid starting position",
        "Wrong character count"
    ],

    alternative:'=TEXTBEFORE(TEXTAFTER(A2,"-"),"-")'
},


// ==========================================================
// LEN
// ==========================================================
{
    title:"LEN",

    keywords:[
        "len","length","count characters",
        "text length","character count",
        "number of characters","data validation",
        "check text length","password length",
        "employee code length","customer id length",
        "string length","text analysis",
        "data cleaning","quality check",
        "validation formula","excel text function",
        "count letters","count digits",
        "text measurement","input validation",
        "master data check","duplicate analysis",
        "report preparation","automation",
        "excel automation","text processing",
        "data quality","advanced excel"
    ],

    aliases:[
        "character counter",
        "text counter",
        "string length checker",
        "text size calculator",
        "length validator",
        "data quality checker",
        "character measurement",
        "input length checker",
        "text analyzer",
        "excel length formula"
    ],

    questions:[
        "How to count characters?",
        "How to find text length?",
        "How to validate code length?",
        "How many characters are in a cell?",
        "How to use LEN function?",
        "How to check employee ID length?",
        "How to validate input data?",
        "How to count letters and numbers?",
        "How to clean incorrect data?",
        "How to create length check?"
    ],

    dataset:"customer_data",

    category:"Text",

    formula:"=LEN(A2)",

    syntax:"LEN(text)",

    explanation:"Returns the number of characters in a text string, including spaces.",

    breakdown:[
        "A2 = Source Text",
        "Counts all characters",
        "Includes spaces"
    ],

    output:"10",

    business:"Used for data validation, ID checks, and identifying incorrect input values.",

    tips:[
        "Combine with IF for validation.",
        "Spaces are counted as characters.",
        "Useful in data cleaning."
    ],

    errors:[
        "#VALUE!",
        "Invalid reference"
    ],

    alternative:"=LEN(TRIM(A2))"
},


// ==========================================================
// FIND
// ==========================================================
{
    title:"FIND",

    keywords:[
        "find","find text","search text",
        "text position","character position",
        "locate text","find symbol",
        "find dash","find space",
        "text location","string search",
        "position finder","delimiter position",
        "split text","data cleaning",
        "text extraction","excel text function",
        "employee code","email extraction",
        "domain extraction","separator finder",
        "report preparation","automation",
        "excel automation","text parsing",
        "dynamic extraction","business data",
        "advanced excel","data analysis"
    ],

    aliases:[
        "text locator",
        "character finder",
        "position finder",
        "string search formula",
        "delimiter finder",
        "text position formula",
        "separator locator",
        "symbol finder",
        "text parser",
        "excel search position"
    ],

    questions:[
        "How to find text position?",
        "How to locate a character?",
        "How to find dash position?",
        "How to split email address?",
        "How to find space location?",
        "How to use FIND function?",
        "How to extract text before symbol?",
        "How to find delimiter position?",
        "How to locate characters?",
        "How to parse text values?"
    ],

    dataset:"employees",

    category:"Text",

    formula:'=FIND("-",A2)',

    syntax:"FIND(find_text,within_text,[start_num])",

    explanation:"Returns the position of one text string inside another text string.",

    breakdown:[
        "\"-\" = Character to Find",
        "A2 = Source Text",
        "Returns character position"
    ],

    output:"6",

    business:"Used for splitting codes, extracting emails, and cleaning structured text data.",

    tips:[
        "FIND is case-sensitive.",
        "Use SEARCH for case-insensitive search.",
        "Combine with LEFT, RIGHT, and MID."
    ],

    errors:[
        "#VALUE!",
        "Text not found",
        "Incorrect search value"
    ],

    alternative:'=SEARCH("-",A2)'
},

// ==========================================================
// SEARCH
// ==========================================================
{
    title:"SEARCH",

    keywords:[
        "search","search text","find text",
        "locate characters","text position",
        "case insensitive find","find word",
        "find symbol","delimiter search",
        "email extraction","text parsing",
        "string search","data cleaning",
        "text extraction","position finder",
        "split text","excel text function",
        "character location","find keyword",
        "business data cleaning","automation",
        "excel automation","report preparation",
        "dynamic extraction","text analysis",
        "advanced excel","string processing",
        "find occurrence","text validation"
    ],

    aliases:[
        "case insensitive find",
        "text locator",
        "keyword finder",
        "string position finder",
        "character search",
        "text scanner",
        "delimiter locator",
        "text detection",
        "search position formula",
        "excel search function"
    ],

    questions:[
        "How to find text position?",
        "How to search inside text?",
        "How to find word in a sentence?",
        "How to locate special characters?",
        "How to extract email parts?",
        "How to use SEARCH function?",
        "How to find delimiter position?",
        "How to check if text exists?",
        "How to split text dynamically?",
        "How to search without case sensitivity?"
    ],

    dataset:"customer",

    category:"Text",

    formula:'=SEARCH("@",A2)',

    syntax:"SEARCH(find_text,within_text,[start_num])",

    explanation:"Finds the position of a text string inside another text string without considering case.",

    breakdown:[
        "\"@\" = Character to Find",
        "A2 = Source Text",
        "Returns character position"
    ],

    output:"8",

    business:"Used for email parsing, text validation, and cleaning imported business data.",

    tips:[
        "SEARCH is not case-sensitive.",
        "Use FIND when case matters.",
        "Combine with LEFT and RIGHT."
    ],

    errors:[
        "#VALUE!",
        "Text not found",
        "Invalid reference"
    ],

    alternative:'=ISNUMBER(SEARCH("@",A2))'
},


// ==========================================================
// SUBSTITUTE
// ==========================================================
{
    title:"SUBSTITUTE",

    keywords:[
        "substitute","replace text","remove text",
        "replace characters","clean data",
        "remove spaces","replace symbols",
        "text correction","data cleaning",
        "change value","fix spelling",
        "replace delimiter","text manipulation",
        "string replacement","excel text function",
        "remove unwanted characters",
        "standardize data","master data cleaning",
        "report preparation","automation",
        "excel automation","text formatting",
        "replace multiple values","data quality",
        "text transformation","advanced excel",
        "normalize text","business data"
    ],

    aliases:[
        "text replacer",
        "character replacement",
        "data cleaner",
        "string cleaner",
        "text correction formula",
        "replace value formula",
        "symbol remover",
        "text standardizer",
        "data cleanup formula",
        "excel replace function"
    ],

    questions:[
        "How to replace text?",
        "How to remove unwanted characters?",
        "How to clean imported data?",
        "How to replace symbols?",
        "How to remove spaces?",
        "How to use SUBSTITUTE function?",
        "How to standardize text?",
        "How to replace specific words?",
        "How to fix inconsistent data?",
        "How to clean reports?"
    ],

    dataset:"customer_data",

    category:"Text",

    formula:'=SUBSTITUTE(A2,"-","")',

    syntax:"SUBSTITUTE(text,old_text,new_text,[instance_num])",

    explanation:"Replaces existing text with new text in a string.",

    breakdown:[
        "A2 = Original Text",
        "\"-\" = Text to Replace",
        "\"\" = Remove Character"
    ],

    output:"ABC123",

    business:"Used to clean IDs, remove symbols, and standardize business data.",

    tips:[
        "Can replace specific occurrence.",
        "Useful for imported files.",
        "Combine with TRIM for cleaning."
    ],

    errors:[
        "#VALUE!",
        "Wrong text reference"
    ],

    alternative:'=REPLACE(A2,4,1,"")'
},


// ==========================================================
// REPLACE
// ==========================================================
{
    title:"REPLACE",

    keywords:[
        "replace","replace characters",
        "change text position","overwrite text",
        "modify string","text editing",
        "remove characters","insert text",
        "data cleaning","text correction",
        "string manipulation","excel text function",
        "mask account number","hide digits",
        "update codes","change employee id",
        "text transformation","automation",
        "excel automation","report cleanup",
        "business data","text formatting",
        "advanced excel","data preparation",
        "identifier modification","text processing"
    ],

    aliases:[
        "text editor formula",
        "character overwrite",
        "string modifier",
        "text updater",
        "position based replacement",
        "masking formula",
        "code modification",
        "text changer",
        "data correction formula",
        "excel replace formula"
    ],

    questions:[
        "How to replace characters by position?",
        "How to hide account digits?",
        "How to modify text?",
        "How to remove characters from middle?",
        "How to use REPLACE function?",
        "How to update codes?",
        "How to mask sensitive numbers?",
        "How to change fixed positions?",
        "How to edit text automatically?",
        "How to format IDs?"
    ],

    dataset:"banking",

    category:"Text",

    formula:'=REPLACE(A2,1,5,"XXXXX")',

    syntax:"REPLACE(old_text,start_num,num_chars,new_text)",

    explanation:"Replaces characters in a text string based on starting position and length.",

    breakdown:[
        "A2 = Original Text",
        "1 = Starting Position",
        "5 = Characters Replaced"
    ],

    output:"XXXXX1234",

    business:"Used for masking account numbers, updating codes, and formatting identifiers.",

    tips:[
        "Position based replacement.",
        "Use SUBSTITUTE for matching text.",
        "Useful for data masking."
    ],

    errors:[
        "#VALUE!",
        "Invalid position"
    ],

    alternative:'=SUBSTITUTE(A2,"12345","XXXXX")'
},


// ==========================================================
// CONCAT
// ==========================================================
{
    title:"CONCAT",

    keywords:[
        "concat","combine text","merge cells",
        "join values","combine columns",
        "merge names","full name formula",
        "text combination","string joining",
        "create description","combine fields",
        "excel text function","data preparation",
        "report formatting","automation",
        "excel automation","dynamic text",
        "business report","text generation",
        "customer name","employee name",
        "merge data","text builder",
        "advanced excel","data analysis"
    ],

    aliases:[
        "text merger",
        "combine cells",
        "join text formula",
        "string combiner",
        "name builder",
        "text creator",
        "value merger",
        "column merger",
        "excel concatenate",
        "dynamic text formula"
    ],

    questions:[
        "How to combine cells?",
        "How to merge first and last name?",
        "How to join text values?",
        "How to create full name?",
        "How to combine columns?",
        "How to use CONCAT function?",
        "How to create descriptions?",
        "How to merge data fields?",
        "How to build text automatically?",
        "How to combine multiple values?"
    ],

    dataset:"employees",

    category:"Text",

    formula:"=CONCAT(A2,B2)",

    syntax:"CONCAT(text1,text2,...)",

    explanation:"Combines multiple text values into one text string.",

    breakdown:[
        "A2 = First Value",
        "B2 = Second Value",
        "Returns combined text"
    ],

    output:"JohnSmith",

    business:"Used for creating names, IDs, descriptions, and automated reports.",

    tips:[
        "Use TEXTJOIN when separators are required.",
        "Modern replacement for CONCATENATE."
    ],

    errors:[
        "#VALUE!"
    ],

    alternative:'=A2&" "&B2'
},


// ==========================================================
// TEXTJOIN
// ==========================================================
{
    title:"TEXTJOIN",

    keywords:[
        "textjoin","combine text with separator",
        "merge multiple cells","join values",
        "comma separated list","create sentence",
        "combine rows","combine columns",
        "dynamic text","report description",
        "employee skills list","customer details",
        "excel text function","automation",
        "text generation","data presentation",
        "business report","dashboard text",
        "string joining","multiple values",
        "ignore blanks","clean output",
        "advanced excel","text builder",
        "dynamic reporting","excel automation"
    ],

    aliases:[
        "text combiner",
        "separator join formula",
        "list creator",
        "sentence builder",
        "multiple cell merger",
        "dynamic CONCAT",
        "comma list generator",
        "text aggregation",
        "value joiner",
        "excel text merge"
    ],

    questions:[
        "How to combine multiple cells with comma?",
        "How to create text list?",
        "How to merge rows into one cell?",
        "How to ignore blanks while joining?",
        "How to create sentence from data?",
        "How to use TEXTJOIN function?",
        "How to combine employee skills?",
        "How to create dynamic descriptions?",
        "How to merge multiple values?",
        "How to create comma separated output?"
    ],

    dataset:"employees",

    category:"Text",

    formula:'=TEXTJOIN(", ",TRUE,A2:C2)',

    syntax:"TEXTJOIN(delimiter,ignore_empty,text1,...)",

    explanation:"Combines multiple text values using a specified separator.",

    breakdown:[
        "\", \" = Separator",
        "TRUE = Ignore Blank Cells",
        "A2:C2 = Values to Combine"
    ],

    output:"Excel, Power BI, SQL",

    business:"Create dynamic descriptions, reports, and combined information fields.",

    tips:[
        "Better than CONCAT for multiple cells.",
        "Can ignore blanks automatically.",
        "Very useful in reporting."
    ],

    errors:[
        "#VALUE!",
        "Invalid delimiter"
    ],

    alternative:"=CONCAT(A2:C2)"
},

// ==========================================================
// TEXTBEFORE
// ==========================================================
{
    title:"TEXTBEFORE",

    keywords:[
        "textbefore","text before delimiter",
        "extract before character","split text",
        "extract first part","get value before",
        "email username","remove suffix",
        "name extraction","code extraction",
        "delimiter split","text parsing",
        "modern excel","excel 365",
        "data cleaning","string extraction",
        "left side extraction","prefix extraction",
        "dynamic text split","report preparation",
        "excel automation","text processing",
        "business data cleaning","advanced excel",
        "extract before dash","extract before symbol",
        "structured text","data transformation"
    ],

    aliases:[
        "extract before symbol",
        "text splitter",
        "prefix extractor",
        "delimiter based extraction",
        "modern left formula",
        "dynamic text extraction",
        "first part extractor",
        "text parser",
        "string separator formula",
        "excel text split"
    ],

    questions:[
        "How to extract text before delimiter?",
        "How to get email username?",
        "How to split text before dash?",
        "How to extract first part of text?",
        "How to use TEXTBEFORE function?",
        "How to remove text after symbol?",
        "How to extract employee code?",
        "How to split names?",
        "How to clean structured text?",
        "How to replace LEFT FIND formula?"
    ],

    dataset:"employees",

    category:"Text",

    formula:'=TEXTBEFORE(A2,"@")',

    syntax:"TEXTBEFORE(text,delimiter,[instance_num])",

    explanation:"Returns the text that appears before a specified delimiter.",

    breakdown:[
        "A2 = Source Text",
        "\"@\" = Separator",
        "Returns text before @"
    ],

    output:"sunil",

    business:"Used for email extraction, employee codes, and structured data cleaning.",

    tips:[
        "Modern replacement for LEFT + FIND.",
        "Available in Excel 365.",
        "Works with multiple delimiters."
    ],

    errors:[
        "#N/A",
        "Delimiter not found",
        "Invalid text"
    ],

    alternative:'=LEFT(A2,FIND("@",A2)-1)'
},


// ==========================================================
// TEXTAFTER
// ==========================================================
{
    title:"TEXTAFTER",

    keywords:[
        "textafter","text after delimiter",
        "extract after character","get value after",
        "split text","extract domain",
        "email domain","suffix extraction",
        "remove prefix","delimiter parsing",
        "text processing","modern excel",
        "excel 365","data cleaning",
        "string extraction","right extraction",
        "dynamic split","business data",
        "report preparation","automation",
        "text transformation","advanced excel",
        "extract after dash","extract after symbol",
        "structured data","data analysis",
        "text parsing"
    ],

    aliases:[
        "extract after symbol",
        "suffix extractor",
        "domain extractor",
        "text parser",
        "modern RIGHT formula",
        "delimiter extractor",
        "second part extractor",
        "dynamic text extraction",
        "string split formula",
        "excel text after"
    ],

    questions:[
        "How to extract text after delimiter?",
        "How to get email domain?",
        "How to split text after dash?",
        "How to remove prefix?",
        "How to use TEXTAFTER function?",
        "How to extract last part of text?",
        "How to parse text values?",
        "How to extract domain name?",
        "How to replace RIGHT FIND?",
        "How to clean imported text?"
    ],

    dataset:"employees",

    category:"Text",

    formula:'=TEXTAFTER(A2,"@")',

    syntax:"TEXTAFTER(text,delimiter,[instance_num])",

    explanation:"Returns the text that appears after a specified delimiter.",

    breakdown:[
        "A2 = Source Text",
        "\"@\" = Separator",
        "Returns text after @"
    ],

    output:"company.com",

    business:"Used for email domain extraction and structured text analysis.",

    tips:[
        "Modern replacement for RIGHT + FIND.",
        "Works with repeated delimiters.",
        "Available in Excel 365."
    ],

    errors:[
        "#N/A",
        "Delimiter missing",
        "Invalid reference"
    ],

    alternative:'=RIGHT(A2,LEN(A2)-FIND("@",A2))'
},


// ==========================================================
// TEXTSPLIT
// ==========================================================
{
    title:"TEXTSPLIT",

    keywords:[
        "textsplit","split text","split column",
        "separate values","delimiter split",
        "comma split","space split",
        "convert text to columns",
        "dynamic split","excel 365",
        "modern excel","data cleaning",
        "import data cleanup","csv split",
        "employee details split",
        "name separation","address split",
        "text parsing","string separation",
        "data transformation","automation",
        "report preparation","advanced excel",
        "replace text to columns",
        "split multiple values",
        "business data cleaning"
    ],

    aliases:[
        "text splitter",
        "dynamic text to columns",
        "delimiter separator",
        "string splitter",
        "data separator",
        "column split formula",
        "excel split function",
        "value separator",
        "text parser",
        "modern split formula"
    ],

    questions:[
        "How to split text into columns?",
        "How to separate comma values?",
        "How to split employee name?",
        "How to convert text to columns formula?",
        "How to split CSV data?",
        "How to use TEXTSPLIT function?",
        "How to separate address fields?",
        "How to split multiple values?",
        "How to clean imported text?",
        "How to divide text automatically?"
    ],

    dataset:"customer",

    category:"Text",

    formula:'=TEXTSPLIT(A2,",")',

    syntax:"TEXTSPLIT(text,col_delimiter,[row_delimiter])",

    explanation:"Splits text into multiple cells using specified delimiters.",

    breakdown:[
        "A2 = Source Text",
        "\",\" = Column Separator",
        "Creates multiple cells"
    ],

    output:"Name | City | State",

    business:"Used for CSV cleaning, address splitting, and preparing datasets.",

    tips:[
        "Replacement for Text to Columns.",
        "Works dynamically.",
        "Available in Excel 365."
    ],

    errors:[
        "#VALUE!",
        "Missing delimiter"
    ],

    alternative:"Data → Text to Columns"
},


// ==========================================================
// TRIM
// ==========================================================
{
    title:"TRIM",

    keywords:[
        "trim","remove spaces","extra spaces",
        "clean text","remove unwanted spaces",
        "data cleaning","text cleanup",
        "duplicate issue fix","match issue",
        "lookup problem","vlookup mismatch",
        "employee data cleaning",
        "customer data cleaning",
        "remove leading spaces",
        "remove trailing spaces",
        "excel text function",
        "master data cleanup",
        "imported data cleaning",
        "report preparation",
        "data quality","automation",
        "excel automation","text formatting",
        "advanced excel"
    ],

    aliases:[
        "space remover",
        "text cleaner",
        "data cleanup formula",
        "blank space remover",
        "string cleaner",
        "lookup fixer",
        "format cleaner",
        "text standardizer",
        "excel cleanup function",
        "space correction"
    ],

    questions:[
        "How to remove extra spaces?",
        "How to clean imported data?",
        "How to fix lookup mismatch?",
        "How to remove spaces before text?",
        "How to use TRIM function?",
        "How to clean employee names?",
        "How to standardize text?",
        "How to remove unwanted blanks?",
        "How to prepare data for analysis?",
        "How to fix duplicate records?"
    ],

    dataset:"employees",

    category:"Text",

    formula:"=TRIM(A2)",

    syntax:"TRIM(text)",

    explanation:"Removes extra spaces from text except single spaces between words.",

    breakdown:[
        "A2 = Source Text",
        "Removes unnecessary spaces",
        "Returns clean text"
    ],

    output:"Sunil Kumar",

    business:"Used before lookup, reporting, and data analysis to clean inconsistent text.",

    tips:[
        "Very useful before VLOOKUP/XLOOKUP.",
        "Does not remove non-printing characters.",
        "Combine with CLEAN for imported data."
    ],

    errors:[
        "#VALUE!"
    ],

    alternative:"=CLEAN(TRIM(A2))"
},


// ==========================================================
// CLEAN
// ==========================================================
{
    title:"CLEAN",

    keywords:[
        "clean","remove non printable characters",
        "data cleaning","remove hidden characters",
        "imported data cleanup","copy paste issue",
        "system export cleaning",
        "excel text function","text cleanup",
        "remove control characters",
        "database cleaning","report preparation",
        "data quality","automation",
        "excel automation","text formatting",
        "hidden symbols","dirty data",
        "character cleanup","advanced excel",
        "master data cleaning","text standardization",
        "remove line breaks","import cleanup"
    ],

    aliases:[
        "hidden character remover",
        "data purifier",
        "text cleaner",
        "export cleanup",
        "control character remover",
        "dirty data cleaner",
        "system data cleaner",
        "text normalization",
        "excel clean function",
        "character cleaner"
    ],

    questions:[
        "How to remove hidden characters?",
        "How to clean copied data?",
        "How to fix imported text?",
        "How to remove line breaks?",
        "How to use CLEAN function?",
        "How to clean system reports?",
        "How to remove unwanted symbols?",
        "How to prepare data for analysis?",
        "How to fix dirty data?",
        "How to clean Excel exports?"
    ],

    dataset:"system_export",

    category:"Text",

    formula:"=CLEAN(A2)",

    syntax:"CLEAN(text)",

    explanation:"Removes non-printable characters from text.",

    breakdown:[
        "A2 = Imported Text",
        "Removes hidden characters",
        "Returns clean value"
    ],

    output:"Clean Employee Name",

    business:"Used to clean exported reports from systems and improve data quality.",

    tips:[
        "Combine with TRIM.",
        "Useful for copied data from websites.",
        "Improves lookup accuracy."
    ],

    errors:[
        "#VALUE!"
    ],

    alternative:"=TRIM(CLEAN(A2))"
},

// ==========================================================
// TODAY
// ==========================================================
{
    title:"TODAY",

    keywords:[
        "today","current date","today date",
        "automatically update date","live date",
        "current day","date formula",
        "dynamic date","daily report date",
        "mis report date","dashboard date",
        "attendance date","leave tracking",
        "expiry date check","due date",
        "deadline calculation","age calculation",
        "date comparison","financial reporting",
        "excel date function","date automation",
        "report refresh date","current working date",
        "business date","system date",
        "real time date","dynamic reporting",
        "advanced excel","date analysis"
    ],

    aliases:[
        "current date formula",
        "live date formula",
        "automatic date",
        "today function",
        "dynamic calendar date",
        "excel current date"
    ],

    questions:[
        "How to get current date?",
        "How to insert today's date automatically?",
        "How to create dynamic date?",
        "How to show report generation date?",
        "How to calculate due date?",
        "How to use TODAY function?",
        "How to compare dates?",
        "How to track expiry date?",
        "How to create daily MIS date?",
        "How to automate date updates?"
    ],

    dataset:"reports",

    category:"Date & Time",

    formula:"=TODAY()",

    syntax:"TODAY()",

    explanation:"Returns the current date and updates automatically when the workbook recalculates.",

    breakdown:[
        "No arguments required",
        "Returns current system date"
    ],

    output:"28-Jul-2026",

    business:"Used in MIS reports, attendance tracking, due date monitoring, and dashboards.",

    tips:[
        "Updates automatically.",
        "Format cell as Date.",
        "Use with YEAR, MONTH, DAY."
    ],

    errors:[
        "Incorrect date format"
    ],

    alternative:"=NOW()"
},


// ==========================================================
// NOW
// ==========================================================
{
    title:"NOW",

    keywords:[
        "now","current date time",
        "live timestamp","current timestamp",
        "date and time","time tracking",
        "system time","real time",
        "report timestamp","refresh time",
        "activity tracking","log time",
        "audit timestamp","process monitoring",
        "workflow tracking","excel date function",
        "automation","MIS timestamp",
        "daily report time","submission time",
        "transaction time","data refresh",
        "business monitoring","time calculation",
        "dashboard timestamp","advanced excel"
    ],

    aliases:[
        "timestamp formula",
        "live date time",
        "current time formula",
        "automatic timestamp",
        "date time function"
    ],

    questions:[
        "How to get current date and time?",
        "How to create timestamp?",
        "How to track report time?",
        "How to show refresh time?",
        "How to use NOW function?",
        "How to record current time?",
        "How to create audit timestamp?",
        "How to monitor process time?",
        "How to track submission time?",
        "How to automate timestamps?"
    ],

    dataset:"audit_log",

    category:"Date & Time",

    formula:"=NOW()",

    syntax:"NOW()",

    explanation:"Returns the current date and time.",

    breakdown:[
        "No arguments required",
        "Returns current date + time"
    ],

    output:"28-Jul-2026 08:30 PM",

    business:"Used for audit logs, process tracking, and automated reporting.",

    tips:[
        "Updates during recalculation.",
        "Use formatting to display only date or time."
    ],

    errors:[
        "Incorrect date formatting"
    ],

    alternative:"=TODAY()"
},


// ==========================================================
// DATE
// ==========================================================
{
    title:"DATE",

    keywords:[
        "date","create date","combine year month day",
        "generate date","build date formula",
        "dynamic date creation","date calculation",
        "financial year date","joining date",
        "expiry date","due date",
        "loan date","transaction date",
        "employee date","report date",
        "excel date function","date automation",
        "date conversion","calendar formula",
        "business date","MIS reporting",
        "period calculation","advanced excel",
        "date analysis","month end date"
    ],

    aliases:[
        "date creator",
        "date generator",
        "calendar formula",
        "date builder",
        "combine date formula"
    ],

    questions:[
        "How to create date from year month day?",
        "How to combine date parts?",
        "How to generate date formula?",
        "How to create financial dates?",
        "How to use DATE function?",
        "How to calculate month end?",
        "How to create joining date?",
        "How to build dynamic dates?",
        "How to convert numbers into date?",
        "How to automate dates?"
    ],

    dataset:"employees",

    category:"Date & Time",

    formula:"=DATE(2026,7,28)",

    syntax:"DATE(year,month,day)",

    explanation:"Creates a valid Excel date from separate year, month, and day values.",

    breakdown:[
        "2026 = Year",
        "7 = Month",
        "28 = Day"
    ],

    output:"28-Jul-2026",

    business:"Used for creating reporting periods, joining dates, expiry dates, and financial calendars.",

    tips:[
        "Useful when dates come from separate columns.",
        "Handles month/year rollover automatically."
    ],

    errors:[
        "#VALUE!",
        "Invalid date values"
    ],

    alternative:"=TODAY()"
},


// ==========================================================
// YEAR
// ==========================================================
{
    title:"YEAR",

    keywords:[
        "year","extract year","get year from date",
        "date analysis","financial year",
        "year wise report","annual report",
        "date breakdown","MIS analysis",
        "trend analysis","employee joining year",
        "transaction year","sales year",
        "calendar year","excel date function",
        "date extraction","reporting",
        "dashboard analysis","period analysis",
        "business intelligence","automation",
        "date transformation","advanced excel",
        "year calculation","data analysis"
    ],

    aliases:[
        "year extractor",
        "date year formula",
        "annual grouping",
        "year finder",
        "calendar year function"
    ],

    questions:[
        "How to extract year from date?",
        "How to create year wise report?",
        "How to find joining year?",
        "How to group data by year?",
        "How to use YEAR function?",
        "How to analyze annual data?",
        "How to get financial year?",
        "How to separate year?",
        "How to create yearly dashboard?",
        "How to calculate year?"
    ],

    dataset:"transactions",

    category:"Date & Time",

    formula:"=YEAR(A2)",

    syntax:"YEAR(serial_number)",

    explanation:"Returns the year from a valid Excel date.",

    breakdown:[
        "A2 = Date Cell",
        "Returns Year Number"
    ],

    output:"2026",

    business:"Used for yearly MIS, trend analysis, and financial reporting.",

    tips:[
        "Combine with MONTH for period analysis.",
        "Works with Pivot reporting."
    ],

    errors:[
        "#VALUE!",
        "Invalid date"
    ],

    alternative:'=TEXT(A2,"YYYY")'
},


// ==========================================================
// MONTH
// ==========================================================
{
    title:"MONTH",

    keywords:[
        "month","extract month","get month from date",
        "monthly report","month wise analysis",
        "MIS month","financial month",
        "date breakdown","sales month",
        "employee joining month","transaction month",
        "period analysis","calendar month",
        "dashboard reporting","excel date function",
        "date extraction","monthly trend",
        "business reporting","automation",
        "date transformation","advanced excel",
        "month calculation","data analysis"
    ],

    aliases:[
        "month extractor",
        "date month formula",
        "monthly grouping",
        "period finder",
        "calendar month function"
    ],

    questions:[
        "How to extract month from date?",
        "How to create monthly report?",
        "How to group data by month?",
        "How to find joining month?",
        "How to use MONTH function?",
        "How to analyze monthly sales?",
        "How to create MIS month?",
        "How to separate month?",
        "How to build monthly dashboard?",
        "How to calculate month?"
    ],

    dataset:"transactions",

    category:"Date & Time",

    formula:"=MONTH(A2)",

    syntax:"MONTH(serial_number)",

    explanation:"Returns the month number from a date.",

    breakdown:[
        "A2 = Date Cell",
        "Returns Month Number"
    ],

    output:"7",

    business:"Used for monthly MIS, trend analysis, and period-based reporting.",

    tips:[
        "Combine with YEAR for monthly trends.",
        "Use TEXT for month names."
    ],

    errors:[
        "#VALUE!",
        "Invalid date"
    ],

    alternative:'=TEXT(A2,"MMM")'
},

// ==========================================================
// DAY
// ==========================================================
{
    title:"DAY",

    keywords:[
        "day","extract day","get day from date",
        "date day formula","day number",
        "daily report","attendance date",
        "transaction day","invoice date",
        "date breakdown","calendar day",
        "employee joining day","date analysis",
        "excel date function","date extraction",
        "daily MIS","period analysis",
        "report automation","business date",
        "calendar calculation","advanced excel",
        "date transformation","day calculation",
        "working day analysis","data analysis"
    ],

    aliases:[
        "day extractor",
        "date day function",
        "calendar day formula",
        "day number finder",
        "date component extractor"
    ],

    questions:[
        "How to extract day from date?",
        "How to get day number?",
        "How to find transaction day?",
        "How to use DAY function?",
        "How to create daily report?",
        "How to separate date parts?",
        "How to calculate day from date?",
        "How to analyze daily data?"
    ],

    dataset:"transactions",

    category:"Date & Time",

    formula:"=DAY(A2)",

    syntax:"DAY(serial_number)",

    explanation:"Returns the day number from a valid Excel date.",

    breakdown:[
        "A2 = Date Cell",
        "Returns Day Number"
    ],

    output:"28",

    business:"Used for daily reporting, attendance tracking, and date analysis.",

    tips:[
        "Combine with MONTH and YEAR for complete date analysis.",
        "Useful for calendar-based reports."
    ],

    errors:[
        "#VALUE!",
        "Invalid date"
    ],

    alternative:'=TEXT(A2,"DD")'
},


// ==========================================================
// DATEDIF
// ==========================================================
{
    title:"DATEDIF",

    keywords:[
        "datedif","date difference","calculate age",
        "calculate duration","difference between dates",
        "years months days","employee age",
        "service period","tenure calculation",
        "experience calculation","leave duration",
        "employment duration","sla days",
        "contract period","loan period",
        "date calculation","excel hidden function",
        "hr analytics","employee lifecycle",
        "business duration","period calculation",
        "time difference","advanced excel",
        "date analysis","working period"
    ],

    aliases:[
        "age calculator",
        "date gap formula",
        "tenure calculator",
        "service duration formula",
        "date interval calculator",
        "period difference"
    ],

    questions:[
        "How to calculate age in Excel?",
        "How to calculate employee experience?",
        "How to find date difference?",
        "How many years between dates?",
        "How to calculate service period?",
        "How to calculate leave duration?",
        "How to use DATEDIF function?",
        "How to calculate SLA period?",
        "How to find months between dates?"
    ],

    dataset:"employees",

    category:"Date & Time",

    formula:'=DATEDIF(B2,TODAY(),"Y")',

    syntax:"DATEDIF(start_date,end_date,unit)",

    explanation:"Calculates the difference between two dates in years, months, or days.",

    breakdown:[
        "B2 = Start Date",
        "TODAY() = Current Date",
        "\"Y\" = Difference in Years"
    ],

    output:"5 Years",

    business:"Used for employee age, experience, tenure, and contract calculations.",

    tips:[
        "Units: Y = Years, M = Months, D = Days.",
        "Useful in HR dashboards."
    ],

    errors:[
        "#NUM!",
        "Start date after end date"
    ],

    alternative:"=YEARFRAC(B2,TODAY())"
},


// ==========================================================
// DAYS
// ==========================================================
{
    title:"DAYS",

    keywords:[
        "days","date difference",
        "number of days between dates",
        "calculate days","duration days",
        "leave days","sla days",
        "payment due days","aging analysis",
        "invoice aging","turnaround time",
        "date calculation","period tracking",
        "business days","excel date function",
        "reporting days","date comparison",
        "transaction ageing","collections report",
        "advanced excel","time analysis",
        "duration calculation","mis reporting"
    ],

    aliases:[
        "date difference formula",
        "day counter",
        "duration calculator",
        "aging calculator",
        "date gap formula"
    ],

    questions:[
        "How to calculate days between dates?",
        "How to find aging days?",
        "How to calculate leave days?",
        "How to calculate SLA days?",
        "How to use DAYS function?",
        "How to find invoice age?",
        "How many days between dates?",
        "How to calculate duration?"
    ],

    dataset:"transactions",

    category:"Date & Time",

    formula:"=DAYS(B2,A2)",

    syntax:"DAYS(end_date,start_date)",

    explanation:"Returns the number of days between two dates.",

    breakdown:[
        "B2 = End Date",
        "A2 = Start Date",
        "Returns Difference in Days"
    ],

    output:"45",

    business:"Used for aging reports, SLA tracking, leave calculation, and turnaround analysis.",

    tips:[
        "Use NETWORKDAYS for working days.",
        "Useful for operational dashboards."
    ],

    errors:[
        "#VALUE!",
        "Invalid date format"
    ],

    alternative:"=B2-A2"
},


// ==========================================================
// EDATE
// ==========================================================
{
    title:"EDATE",

    keywords:[
        "edate","add months","subtract months",
        "future date","previous date",
        "maturity date","expiry date",
        "renewal date","contract renewal",
        "loan maturity","subscription expiry",
        "date calculation","month calculation",
        "financial date","investment period",
        "excel date function","period extension",
        "calendar calculation","business date",
        "advanced excel","date automation",
        "dynamic dates","reporting period"
    ],

    aliases:[
        "month addition formula",
        "future date calculator",
        "renewal date formula",
        "maturity calculator",
        "date shift formula"
    ],

    questions:[
        "How to add months to date?",
        "How to calculate maturity date?",
        "How to add 12 months?",
        "How to calculate renewal date?",
        "How to use EDATE function?",
        "How to find previous month date?",
        "How to create future dates?"
    ],

    dataset:"contracts",

    category:"Date & Time",

    formula:"=EDATE(A2,12)",

    syntax:"EDATE(start_date,months)",

    explanation:"Returns a date before or after a specified number of months.",

    breakdown:[
        "A2 = Starting Date",
        "12 = Months to Add",
        "Returns Future Date"
    ],

    output:"28-Jul-2027",

    business:"Used for contracts, renewals, maturity dates, and financial planning.",

    tips:[
        "Positive months adds future dates.",
        "Negative months gives past dates."
    ],

    errors:[
        "#VALUE!",
        "Invalid date"
    ],

    alternative:"=DATE(YEAR(A2)+1,MONTH(A2),DAY(A2))"
},


// ==========================================================
// EOMONTH
// ==========================================================
{
    title:"EOMONTH",

    keywords:[
        "eomonth","end of month",
        "month end date","month closing",
        "financial month end","MIS closing",
        "reporting period","last date month",
        "month end calculation","accounting date",
        "balance sheet date","period end",
        "monthly report","quarter end",
        "year end","excel date function",
        "calendar calculation","business reporting",
        "automation","finance reporting",
        "advanced excel","date analysis",
        "month closing process","MIS automation"
    ],

    aliases:[
        "month end formula",
        "closing date formula",
        "last day month",
        "period end calculator",
        "financial closing date"
    ],

    questions:[
        "How to find month end date?",
        "How to calculate last date of month?",
        "How to create month closing report?",
        "How to use EOMONTH function?",
        "How to find quarter end date?",
        "How to calculate financial period end?",
        "How to create MIS month end?"
    ],

    dataset:"finance",

    category:"Date & Time",

    formula:"=EOMONTH(A2,0)",

    syntax:"EOMONTH(start_date,months)",

    explanation:"Returns the last day of a month before or after a specified number of months.",

    breakdown:[
        "A2 = Reference Date",
        "0 = Same Month",
        "Returns Month End Date"
    ],

    output:"31-Jul-2026",

    business:"Used for month-end closing, MIS reporting, accounting, and financial analysis.",

    tips:[
        "Very useful for monthly dashboards.",
        "Use months=1 for next month end."
    ],

    errors:[
        "#VALUE!",
        "Invalid date"
    ],

    alternative:'=DATE(YEAR(A2),MONTH(A2)+1,0)'
},
// ==========================================================
// WEEKDAY
// ==========================================================
{
    title:"WEEKDAY",

    keywords:[
        "weekday","day name","day number",
        "find day of week","monday tuesday",
        "date to day","calendar analysis",
        "attendance analysis","working day",
        "week calculation","date function",
        "excel weekday formula","daily report",
        "transaction day","business day",
        "date breakdown","calendar reporting",
        "schedule planning","employee attendance",
        "leave analysis","automation",
        "dashboard date","advanced excel",
        "date intelligence","time analysis"
    ],

    aliases:[
        "day finder",
        "day of week formula",
        "calendar day extractor",
        "date weekday calculator",
        "week day number"
    ],

    questions:[
        "How to find day from date?",
        "How to know Monday or Sunday?",
        "How to calculate weekday?",
        "How to identify working days?",
        "How to use WEEKDAY function?",
        "How to create attendance analysis?",
        "How to check transaction day?",
        "How to convert date into weekday?"
    ],

    dataset:"attendance",

    category:"Date & Time",

    formula:"=WEEKDAY(A2,2)",

    syntax:"WEEKDAY(serial_number,[return_type])",

    explanation:"Returns the day number of a week from a date.",

    breakdown:[
        "A2 = Date",
        "2 = Monday as first day",
        "Returns weekday number"
    ],

    output:"1 (Monday)",

    business:"Used for attendance reports, scheduling, and working day analysis.",

    tips:[
        "Return type 2 makes Monday = 1.",
        "Combine with TEXT for day names."
    ],

    errors:[
        "#VALUE!",
        "Invalid date"
    ],

    alternative:'=TEXT(A2,"DDD")'
},


// ==========================================================
// WEEKNUM
// ==========================================================
{
    title:"WEEKNUM",

    keywords:[
        "weeknum","week number",
        "calculate week","week wise report",
        "weekly MIS","calendar week",
        "week analysis","date grouping",
        "weekly dashboard","sales week",
        "transaction week","performance week",
        "excel date function","period analysis",
        "business reporting","calendar calculation",
        "employee productivity","weekly trend",
        "automation","advanced excel",
        "date intelligence","reporting period",
        "week calculation","time analysis"
    ],

    aliases:[
        "week calculator",
        "week grouping formula",
        "calendar week formula",
        "weekly report formula",
        "week identifier"
    ],

    questions:[
        "How to find week number?",
        "How to create weekly report?",
        "How to group dates by week?",
        "How to use WEEKNUM function?",
        "How to analyze weekly sales?",
        "How to create weekly MIS?",
        "How to identify calendar week?"
    ],

    dataset:"sales",

    category:"Date & Time",

    formula:"=WEEKNUM(A2,2)",

    syntax:"WEEKNUM(serial_number,[return_type])",

    explanation:"Returns the week number of a date.",

    breakdown:[
        "A2 = Date",
        "2 = Week starts Monday",
        "Returns week number"
    ],

    output:"Week 30",

    business:"Used for weekly MIS, productivity tracking, and trend analysis.",

    tips:[
        "Useful for weekly dashboards.",
        "Use with YEAR for yearly weekly analysis."
    ],

    errors:[
        "#VALUE!",
        "Invalid date"
    ],

    alternative:'="Week "&WEEKNUM(A2,2)'
},


// ==========================================================
// NETWORKDAYS
// ==========================================================
{
    title:"NETWORKDAYS",

    keywords:[
        "networkdays","working days",
        "business days","exclude weekends",
        "exclude holidays","sla days",
        "turnaround time","tat calculation",
        "leave calculation","employee leave",
        "project duration","processing days",
        "bank working days","operational days",
        "date difference","excel date function",
        "holiday calculation","compliance sla",
        "service period","working duration",
        "mis reporting","automation",
        "advanced excel","time analysis",
        "business calendar"
    ],

    aliases:[
        "working day calculator",
        "business day formula",
        "SLA calculator",
        "TAT calculator",
        "office day counter",
        "holiday excluded days"
    ],

    questions:[
        "How to calculate working days?",
        "How to exclude weekends?",
        "How to calculate SLA days?",
        "How to calculate leave days?",
        "How to ignore holidays?",
        "How to use NETWORKDAYS?",
        "How to calculate turnaround time?",
        "How to count office days?"
    ],

    dataset:"operations",

    category:"Date & Time",

    formula:"=NETWORKDAYS(A2,B2)",

    syntax:"NETWORKDAYS(start_date,end_date,[holidays])",

    explanation:"Returns the number of working days between two dates excluding weekends and optional holidays.",

    breakdown:[
        "A2 = Start Date",
        "B2 = End Date",
        "Excludes Saturday and Sunday"
    ],

    output:"22 Working Days",

    business:"Used for SLA tracking, employee leave, and operational performance reports.",

    tips:[
        "Add holiday range for accurate calculation.",
        "Very useful for compliance reporting."
    ],

    errors:[
        "#VALUE!",
        "Invalid dates"
    ],

    alternative:"=DAYS(B2,A2)"
},


// ==========================================================
// NETWORKDAYS.INTL
// ==========================================================
{
    title:"NETWORKDAYS.INTL",

    keywords:[
        "networkdays intl","custom working days",
        "different weekend","shift calculation",
        "custom weekend","international calendar",
        "working days formula","exclude weekends",
        "saturday working","friday saturday weekend",
        "global business days","sla calculation",
        "holiday exclusion","employee shift",
        "date calculation","advanced excel",
        "operational reporting","compliance tracking",
        "tat calculation","business calendar",
        "automation","time analysis",
        "regional calendar","working schedule"
    ],

    aliases:[
        "custom business days",
        "shift working days",
        "international working calendar",
        "custom weekend formula",
        "flexible SLA calculator"
    ],

    questions:[
        "How to calculate custom working days?",
        "How to change weekends?",
        "How to calculate Friday Saturday weekend?",
        "How to use NETWORKDAYS.INTL?",
        "How to handle global holidays?",
        "How to calculate shift working days?"
    ],

    dataset:"global_operations",

    category:"Date & Time",

    formula:'=NETWORKDAYS.INTL(A2,B2,"0000110")',

    syntax:"NETWORKDAYS.INTL(start_date,end_date,[weekend],[holidays])",

    explanation:"Calculates working days using custom weekend patterns.",

    breakdown:[
        "A2 = Start Date",
        "B2 = End Date",
        "0000110 = Custom Weekend Pattern"
    ],

    output:"20 Working Days",

    business:"Used for global teams with different weekend schedules.",

    tips:[
        "Useful for multinational operations.",
        "Supports custom working patterns."
    ],

    errors:[
        "#VALUE!",
        "Invalid weekend code"
    ],

    alternative:"=NETWORKDAYS(A2,B2)"
},


// ==========================================================
// WORKDAY
// ==========================================================
{
    title:"WORKDAY",

    keywords:[
        "workday","future working date",
        "add working days","calculate deadline",
        "due date","sla deadline",
        "project completion date",
        "exclude weekends","exclude holidays",
        "business date","processing date",
        "expected completion","task deadline",
        "excel date function","operational planning",
        "workflow management","compliance date",
        "turnaround time","automation",
        "advanced excel","date calculation",
        "business calendar","project tracking"
    ],

    aliases:[
        "deadline calculator",
        "working date formula",
        "business date generator",
        "future date calculator",
        "SLA deadline formula"
    ],

    questions:[
        "How to add working days?",
        "How to calculate deadline date?",
        "How to exclude weekends?",
        "How to find project completion date?",
        "How to use WORKDAY function?",
        "How to calculate SLA due date?",
        "How to add business days?"
    ],

    dataset:"projects",

    category:"Date & Time",

    formula:"=WORKDAY(A2,10)",

    syntax:"WORKDAY(start_date,days,[holidays])",

    explanation:"Returns a future or past working date excluding weekends and holidays.",

    breakdown:[
        "A2 = Start Date",
        "10 = Working Days To Add",
        "Returns Due Date"
    ],

    output:"11-Aug-2026",

    business:"Used for SLA deadlines, project planning, and operational tracking.",

    tips:[
        "Use NETWORKDAYS to count days.",
        "Use WORKDAY to calculate future dates."
    ],

    errors:[
        "#VALUE!",
        "Invalid date"
    ],

    alternative:"=A2+10"
},

// ==========================================================
// SUM
// ==========================================================
{
    title:"SUM",

    keywords:[
        "sum","total","add numbers",
        "calculate total","grand total",
        "addition formula","amount total",
        "sales total","expense total",
        "salary total","revenue calculation",
        "financial summary","mis total",
        "dashboard calculation","numeric addition",
        "excel math function","data analysis",
        "report preparation","automation",
        "business calculation","kpi total",
        "monthly total","yearly total",
        "advanced excel","finance reporting",
        "amount aggregation","value addition"
    ],

    aliases:[
        "total formula",
        "addition formula",
        "grand total calculator",
        "amount aggregator",
        "excel add function"
    ],

    questions:[
        "How to calculate total?",
        "How to add numbers in Excel?",
        "How to create grand total?",
        "How to calculate sales total?",
        "How to sum multiple cells?",
        "How to use SUM function?",
        "How to calculate expense total?",
        "How to create MIS total?"
    ],

    dataset:"sales",

    category:"Math",

    formula:"=SUM(B2:B10)",

    syntax:"SUM(number1,[number2],...)",

    explanation:"Adds all numbers in a selected range and returns the total value.",

    breakdown:[
        "B2:B10 = Range of Values",
        "SUM adds all numbers",
        "Returns total amount"
    ],

    output:"125000",

    business:"Used for calculating sales, expenses, salary totals, and financial summaries.",

    tips:[
        "Works with multiple ranges.",
        "Use AutoSum for quick totals.",
        "Ignores text values."
    ],

    errors:[
        "#VALUE!",
        "Wrong range selection"
    ],

    alternative:"=B2+B3+B4"
},


// ==========================================================
// SUMIF
// ==========================================================
{
    title:"SUMIF",

    keywords:[
        "sumif","conditional sum",
        "sum based on condition",
        "total with criteria",
        "category wise total",
        "employee wise salary",
        "department total",
        "sales by region",
        "expense category",
        "amount calculation",
        "conditional calculation",
        "excel criteria formula",
        "mis reporting",
        "dashboard formula",
        "business analysis",
        "financial analysis",
        "data summarization",
        "automation",
        "advanced excel",
        "report preparation",
        "kpi calculation",
        "performance analysis",
        "filter based total"
    ],

    aliases:[
        "criteria total formula",
        "conditional addition",
        "category sum formula",
        "single condition sum",
        "filtered total"
    ],

    questions:[
        "How to sum based on condition?",
        "How to calculate department total?",
        "How to total sales by region?",
        "How to calculate employee salary total?",
        "How to use SUMIF?",
        "How to add values with criteria?",
        "How to create conditional total?"
    ],

    dataset:"employees",

    category:"Math",

    formula:'=SUMIF(B2:B10,"IT",C2:C10)',

    syntax:"SUMIF(range,criteria,sum_range)",

    explanation:"Adds values that match a single condition.",

    breakdown:[
        "B2:B10 = Condition Range",
        "\"IT\" = Matching Criteria",
        "C2:C10 = Values to Add"
    ],

    output:"350000",

    business:"Used for department-wise totals, category analysis, and filtered MIS reports.",

    tips:[
        "Use SUMIFS for multiple conditions.",
        "Criteria can be text, numbers, or expressions."
    ],

    errors:[
        "#VALUE!",
        "Range mismatch"
    ],

    alternative:'=SUMIFS(C:C,B:B,"IT")'
},


// ==========================================================
// SUMIFS
// ==========================================================
{
    title:"SUMIFS",

    keywords:[
        "sumifs","multiple condition sum",
        "advanced sumif",
        "conditional total",
        "multiple criteria",
        "sales by month",
        "department and region total",
        "employee salary analysis",
        "finance reporting",
        "complex calculation",
        "dashboard formula",
        "mis automation",
        "business intelligence",
        "data analysis",
        "advanced excel",
        "criteria based sum",
        "filtered aggregation",
        "multiple filters",
        "reporting automation",
        "performance report",
        "financial analysis",
        "kpi calculation",
        "data summarization",
        "conditional aggregation"
    ],

    aliases:[
        "multi condition sum",
        "advanced conditional total",
        "multiple filter sum",
        "criteria aggregation",
        "complex sum formula"
    ],

    questions:[
        "How to sum with multiple conditions?",
        "How to calculate monthly sales?",
        "How to create department wise region total?",
        "How to use SUMIFS?",
        "How to add multiple criteria?",
        "How to create advanced MIS reports?",
        "How to calculate filtered totals?"
    ],

    dataset:"sales",

    category:"Math",

    formula:'=SUMIFS(D:D,B:B,"North",C:C,"Laptop")',

    syntax:"SUMIFS(sum_range,criteria_range1,criteria1,...)",

    explanation:"Adds values that meet multiple conditions.",

    breakdown:[
        "D:D = Amount Column",
        "B:B = Region",
        "C:C = Product"
    ],

    output:"850000",

    business:"Used for advanced MIS, sales dashboards, and multi-condition analysis.",

    tips:[
        "Most powerful replacement for multiple SUMIF formulas.",
        "Works well with Pivot dashboards."
    ],

    errors:[
        "#VALUE!",
        "Incorrect range size"
    ],

    alternative:'=SUMPRODUCT((B:B="North")*(C:C="Laptop")*D:D)'
},


// ==========================================================
// COUNT
// ==========================================================
{
    title:"COUNT",

    keywords:[
        "count","count numbers",
        "number of records",
        "count transactions",
        "count employees",
        "numeric count",
        "data counting",
        "report summary",
        "kpi count",
        "dashboard metric",
        "transaction volume",
        "excel statistical function",
        "analysis formula",
        "data validation",
        "record counting",
        "business reporting",
        "automation",
        "advanced excel",
        "performance metrics",
        "monthly count",
        "yearly count",
        "numeric analysis"
    ],

    aliases:[
        "number counter",
        "numeric record counter",
        "value counter",
        "transaction counter",
        "excel count formula"
    ],

    questions:[
        "How to count numbers?",
        "How to count transactions?",
        "How many numeric values exist?",
        "How to use COUNT function?",
        "How to count records?",
        "How to create KPI count?"
    ],

    dataset:"transactions",

    category:"Math",

    formula:"=COUNT(A2:A100)",

    syntax:"COUNT(value1,[value2],...)",

    explanation:"Counts the number of cells containing numeric values.",

    breakdown:[
        "A2:A100 = Data Range",
        "Counts only numbers",
        "Ignores text"
    ],

    output:"250",

    business:"Used for transaction counts, KPI reports, and numeric data analysis.",

    tips:[
        "Use COUNTA for text + numbers.",
        "Use COUNTIF for conditions."
    ],

    errors:[
        "#VALUE!",
        "Invalid range"
    ],

    alternative:"=COUNTA(A2:A100)"
},


// ==========================================================
// COUNTA
// ==========================================================
{
    title:"COUNTA",

    keywords:[
        "counta","count records",
        "count non blank",
        "number of entries",
        "filled cells",
        "data completeness",
        "employee count",
        "customer count",
        "record count",
        "missing data check",
        "data quality",
        "excel statistical function",
        "dashboard metrics",
        "reporting",
        "automation",
        "data validation",
        "business analysis",
        "master data",
        "table size",
        "advanced excel",
        "record monitoring",
        "input tracking"
    ],

    aliases:[
        "non blank counter",
        "record counter",
        "entry counter",
        "filled cell counter",
        "data volume calculator"
    ],

    questions:[
        "How to count filled cells?",
        "How many records are available?",
        "How to count employees?",
        "How to check data completeness?",
        "How to use COUNTA function?",
        "How to count text and numbers?",
        "How to count non blank cells?"
    ],

    dataset:"employees",

    category:"Math",

    formula:"=COUNTA(A2:A100)",

    syntax:"COUNTA(value1,[value2],...)",

    explanation:"Counts all non-empty cells including text, numbers, and errors.",

    breakdown:[
        "A2:A100 = Data Range",
        "Counts filled cells",
        "Ignores blanks"
    ],

    output:"98",

    business:"Used for employee count, data completeness checks, and reporting metrics.",

    tips:[
        "Useful for checking uploaded data.",
        "Combine with COUNTBLANK for quality checks."
    ],

    errors:[
        "#VALUE!"
    ],

    alternative:"=ROWS(A2:A100)-COUNTBLANK(A2:A100)"
},

// ==========================================================
// COUNTIF
// ==========================================================
{
    title:"COUNTIF",

    keywords:[
        "countif","count with condition",
        "conditional count","count matching values",
        "count employees","department count",
        "category count","status count",
        "attendance count","sales count",
        "error count","duplicate count",
        "excel criteria formula",
        "dashboard metric","mis reporting",
        "data analysis","business reporting",
        "automation","advanced excel",
        "performance analysis",
        "filtered count","single criteria count",
        "record counting","kpi calculation",
        "data validation"
    ],

    aliases:[
        "criteria counter",
        "conditional counter",
        "filtered record count",
        "single condition count",
        "category counter"
    ],

    questions:[
        "How to count based on condition?",
        "How many employees in a department?",
        "How to count status values?",
        "How to count matching records?",
        "How to use COUNTIF?",
        "How to count sales by category?",
        "How to count errors in report?"
    ],

    dataset:"employees",

    category:"Math",

    formula:'=COUNTIF(B2:B100,"IT")',

    syntax:"COUNTIF(range,criteria)",

    explanation:"Counts cells that meet a specific condition.",

    breakdown:[
        "B2:B100 = Condition Range",
        "\"IT\" = Matching Criteria",
        "Returns Matching Count"
    ],

    output:"25",

    business:"Used for department count, status tracking, and KPI reporting.",

    tips:[
        "Use COUNTIFS for multiple conditions.",
        "Supports wildcards like * and ?."
    ],

    errors:[
        "#VALUE!",
        "Invalid range"
    ],

    alternative:'=COUNTIFS(B:B,"IT")'
},


// ==========================================================
// COUNTIFS
// ==========================================================
{
    title:"COUNTIFS",

    keywords:[
        "countifs","multiple condition count",
        "advanced countif",
        "count multiple criteria",
        "employee analysis",
        "sales count",
        "region wise count",
        "status and category count",
        "complex reporting",
        "dashboard formula",
        "multiple filter count",
        "business intelligence",
        "data analysis",
        "advanced excel",
        "kpi calculation",
        "conditional counting",
        "record analysis",
        "performance reporting",
        "automation",
        "mis dashboard",
        "filtered records",
        "multi criteria formula",
        "operational reporting"
    ],

    aliases:[
        "multi condition counter",
        "advanced record count",
        "multiple filter count",
        "criteria based count",
        "complex COUNTIF"
    ],

    questions:[
        "How to count with multiple conditions?",
        "How many employees match criteria?",
        "How to count sales by region and product?",
        "How to use COUNTIFS?",
        "How to create advanced KPI count?",
        "How to count multiple filters?"
    ],

    dataset:"sales",

    category:"Math",

    formula:'=COUNTIFS(B:B,"North",C:C,"Laptop")',

    syntax:"COUNTIFS(criteria_range1,criteria1,...)",

    explanation:"Counts records that satisfy multiple conditions.",

    breakdown:[
        "B:B = Region Column",
        "North = First Condition",
        "C:C = Product Condition"
    ],

    output:"45",

    business:"Used for advanced dashboards, MIS reports, and multi-filter analysis.",

    tips:[
        "All conditions must be TRUE.",
        "Works well with dashboard slicers."
    ],

    errors:[
        "#VALUE!",
        "Range size mismatch"
    ],

    alternative:'=SUMPRODUCT((B:B="North")*(C:C="Laptop"))'
},


// ==========================================================
// AVERAGE
// ==========================================================
{
    title:"AVERAGE",

    keywords:[
        "average","mean",
        "calculate average",
        "average salary",
        "average sales",
        "average score",
        "performance average",
        "monthly average",
        "financial average",
        "excel statistical function",
        "data analysis",
        "trend analysis",
        "dashboard KPI",
        "business reporting",
        "numeric analysis",
        "employee analytics",
        "automation",
        "advanced excel",
        "calculation formula",
        "summary statistics",
        "report preparation",
        "MIS analysis"
    ],

    aliases:[
        "mean calculator",
        "numeric average",
        "average formula",
        "value average",
        "statistical average"
    ],

    questions:[
        "How to calculate average?",
        "How to find average salary?",
        "How to calculate average sales?",
        "How to use AVERAGE function?",
        "How to find average performance?",
        "How to create average KPI?"
    ],

    dataset:"employees",

    category:"Math",

    formula:"=AVERAGE(C2:C100)",

    syntax:"AVERAGE(number1,[number2],...)",

    explanation:"Returns the arithmetic average of selected numeric values.",

    breakdown:[
        "C2:C100 = Numeric Range",
        "Adds values",
        "Divides by count"
    ],

    output:"65000",

    business:"Used for average salary, sales analysis, and performance metrics.",

    tips:[
        "Ignores blank cells.",
        "Use AVERAGEIF for conditions."
    ],

    errors:[
        "#DIV/0!",
        "No numeric values"
    ],

    alternative:"=SUM(C2:C100)/COUNT(C2:C100)"
},


// ==========================================================
// AVERAGEIF
// ==========================================================
{
    title:"AVERAGEIF",

    keywords:[
        "averageif","conditional average",
        "average with condition",
        "average salary department",
        "average sales category",
        "employee performance average",
        "criteria average",
        "filtered average",
        "department analysis",
        "region average",
        "excel statistical formula",
        "dashboard calculation",
        "MIS reporting",
        "business analysis",
        "data analytics",
        "automation",
        "advanced excel",
        "performance metrics",
        "single condition average",
        "conditional calculation",
        "KPI average",
        "report automation"
    ],

    aliases:[
        "criteria average",
        "filtered average formula",
        "conditional mean",
        "single criteria average",
        "category average"
    ],

    questions:[
        "How to calculate average by condition?",
        "How to find department average salary?",
        "How to calculate average sales?",
        "How to use AVERAGEIF?",
        "How to find average based on category?",
        "How to create KPI average?"
    ],

    dataset:"employees",

    category:"Math",

    formula:'=AVERAGEIF(B:B,"IT",C:C)',

    syntax:"AVERAGEIF(range,criteria,average_range)",

    explanation:"Calculates average values based on one condition.",

    breakdown:[
        "B:B = Criteria Column",
        "\"IT\" = Condition",
        "C:C = Average Column"
    ],

    output:"72000",

    business:"Used for department-wise salary, category analysis, and performance reports.",

    tips:[
        "Use AVERAGEIFS for multiple conditions.",
        "Ignores text values."
    ],

    errors:[
        "#DIV/0!",
        "No matching records"
    ],

    alternative:'=AVERAGEIFS(C:C,B:B,"IT")'
},


// ==========================================================
// AVERAGEIFS
// ==========================================================
{
    title:"AVERAGEIFS",

    keywords:[
        "averageifs","multiple condition average",
        "advanced averageif",
        "average multiple criteria",
        "department region average",
        "sales performance average",
        "salary analysis",
        "complex KPI calculation",
        "filtered average",
        "business analytics",
        "dashboard formula",
        "MIS reporting",
        "data analysis",
        "advanced excel",
        "performance measurement",
        "multi criteria formula",
        "conditional statistics",
        "automation",
        "report preparation",
        "employee analytics",
        "financial analysis",
        "trend analysis",
        "business intelligence"
    ],

    aliases:[
        "multi criteria average",
        "advanced conditional average",
        "multiple filter average",
        "complex average formula",
        "criteria based average"
    ],

    questions:[
        "How to calculate average with multiple conditions?",
        "How to find average salary by department and location?",
        "How to calculate sales average by region?",
        "How to use AVERAGEIFS?",
        "How to create advanced KPI averages?",
        "How to analyze multiple criteria?"
    ],

    dataset:"employees",

    category:"Math",

    formula:'=AVERAGEIFS(C:C,B:B,"IT",D:D,"Mumbai")',

    syntax:"AVERAGEIFS(average_range,criteria_range1,criteria1,...)",

    explanation:"Calculates average values that satisfy multiple conditions.",

    breakdown:[
        "C:C = Salary Column",
        "B:B = Department",
        "D:D = Location"
    ],

    output:"78000",

    business:"Used for advanced HR analytics, financial reports, and performance dashboards.",

    tips:[
        "All conditions must match.",
        "Useful for complex MIS reporting."
    ],

    errors:[
        "#DIV/0!",
        "No matching records"
    ],

    alternative:'=SUMIFS(C:C,B:B,"IT",D:D,"Mumbai")/COUNTIFS(B:B,"IT",D:D,"Mumbai")'
},

// ==========================================================
// MAX
// ==========================================================
{
    title:"MAX",

    keywords:[
        "max","maximum value",
        "highest value","largest number",
        "top value","highest salary",
        "maximum sales","peak value",
        "highest transaction",
        "best performer",
        "maximum amount",
        "excel statistical function",
        "data analysis",
        "dashboard KPI",
        "ranking analysis",
        "performance report",
        "financial analysis",
        "MIS report",
        "business analytics",
        "automation",
        "advanced excel",
        "value comparison",
        "highest score",
        "trend analysis"
    ],

    aliases:[
        "highest number formula",
        "maximum finder",
        "top value calculator",
        "largest value function",
        "peak value finder"
    ],

    questions:[
        "How to find highest value?",
        "How to find maximum salary?",
        "How to get highest sales?",
        "How to identify top performer?",
        "How to use MAX function?",
        "How to find largest number?",
        "How to find peak transaction?"
    ],

    dataset:"employees",

    category:"Math",

    formula:"=MAX(C2:C100)",

    syntax:"MAX(number1,[number2],...)",

    explanation:"Returns the largest value from a range of numbers.",

    breakdown:[
        "C2:C100 = Numeric Range",
        "Checks all values",
        "Returns highest value"
    ],

    output:"150000",

    business:"Used for highest salary, maximum sales, and peak performance analysis.",

    tips:[
        "Combine with INDEX MATCH to find related details.",
        "Ignores text values."
    ],

    errors:[
        "#VALUE!",
        "No numeric values"
    ],

    alternative:"=LARGE(C2:C100,1)"
},


// ==========================================================
// MIN
// ==========================================================
{
    title:"MIN",

    keywords:[
        "min","minimum value",
        "lowest value","smallest number",
        "lowest salary","minimum sales",
        "least amount","minimum transaction",
        "best cost","lowest price",
        "excel statistical function",
        "data analysis",
        "financial analysis",
        "dashboard KPI",
        "comparison formula",
        "performance report",
        "MIS reporting",
        "automation",
        "advanced excel",
        "value analysis",
        "trend analysis",
        "business reporting",
        "cost analysis"
    ],

    aliases:[
        "lowest value formula",
        "smallest number finder",
        "minimum calculator",
        "bottom value function",
        "least value formula"
    ],

    questions:[
        "How to find lowest value?",
        "How to find minimum salary?",
        "How to get lowest sales?",
        "How to find smallest number?",
        "How to use MIN function?",
        "How to identify minimum cost?"
    ],

    dataset:"expenses",

    category:"Math",

    formula:"=MIN(B2:B100)",

    syntax:"MIN(number1,[number2],...)",

    explanation:"Returns the smallest value from a range of numbers.",

    breakdown:[
        "B2:B100 = Numeric Range",
        "Checks all values",
        "Returns lowest value"
    ],

    output:"500",

    business:"Used for minimum cost, lowest transaction, and pricing analysis.",

    tips:[
        "Useful with conditional functions.",
        "Ignores blank cells."
    ],

    errors:[
        "#VALUE!",
        "No numeric values"
    ],

    alternative:"=SMALL(B2:B100,1)"
},


// ==========================================================
// LARGE
// ==========================================================
{
    title:"LARGE",

    keywords:[
        "large","nth largest",
        "top values","ranking",
        "top salary","highest sales",
        "top 5 performers",
        "largest transaction",
        "ranking analysis",
        "performance ranking",
        "dashboard ranking",
        "excel statistical function",
        "data analysis",
        "MIS report",
        "business intelligence",
        "leaderboard",
        "employee ranking",
        "sales ranking",
        "advanced excel",
        "automation",
        "comparison analysis",
        "top records"
    ],

    aliases:[
        "top n formula",
        "rank value finder",
        "largest nth value",
        "top performer formula",
        "ranking function"
    ],

    questions:[
        "How to find top 5 values?",
        "How to find second highest salary?",
        "How to rank sales?",
        "How to use LARGE function?",
        "How to get nth largest value?",
        "How to create leaderboard?"
    ],

    dataset:"sales",

    category:"Math",

    formula:"=LARGE(B2:B100,5)",

    syntax:"LARGE(array,k)",

    explanation:"Returns the k-th largest value from a data range.",

    breakdown:[
        "B2:B100 = Data Range",
        "5 = Position",
        "Returns 5th highest value"
    ],

    output:"95000",

    business:"Used for top performer reports, ranking dashboards, and analytics.",

    tips:[
        "Combine with INDEX MATCH to return names.",
        "Useful for Top N reports."
    ],

    errors:[
        "#NUM!",
        "Invalid k value"
    ],

    alternative:"=SORT(B2:B100,,-1)"
},


// ==========================================================
// SMALL
// ==========================================================
{
    title:"SMALL",

    keywords:[
        "small","nth smallest",
        "lowest ranking","bottom values",
        "minimum ranking",
        "lowest cost",
        "smallest transaction",
        "bottom performer",
        "ranking analysis",
        "data analysis",
        "excel statistical function",
        "comparison formula",
        "MIS report",
        "financial analysis",
        "cost analysis",
        "performance analysis",
        "advanced excel",
        "automation",
        "value ranking",
        "top bottom analysis",
        "business reporting",
        "analytics"
    ],

    aliases:[
        "bottom n formula",
        "lowest rank formula",
        "smallest nth value",
        "minimum ranking function",
        "bottom value finder"
    ],

    questions:[
        "How to find lowest 5 values?",
        "How to find second lowest value?",
        "How to rank lowest items?",
        "How to use SMALL function?",
        "How to identify minimum records?",
        "How to create bottom analysis?"
    ],

    dataset:"expenses",

    category:"Math",

    formula:"=SMALL(B2:B100,5)",

    syntax:"SMALL(array,k)",

    explanation:"Returns the k-th smallest value from a range.",

    breakdown:[
        "B2:B100 = Data Range",
        "5 = Position",
        "Returns 5th lowest value"
    ],

    output:"2500",

    business:"Used for cost analysis, bottom rankings, and exception reporting.",

    tips:[
        "Useful for identifying low values.",
        "Combine with INDEX MATCH for details."
    ],

    errors:[
        "#NUM!",
        "Invalid k value"
    ],

    alternative:"=SORT(B2:B100,,1)"
},


// ==========================================================
// ROUND
// ==========================================================
{
    title:"ROUND",

    keywords:[
        "round","round number",
        "decimal rounding",
        "remove decimals",
        "financial rounding",
        "amount formatting",
        "currency calculation",
        "invoice rounding",
        "percentage rounding",
        "salary calculation",
        "excel math function",
        "number formatting",
        "report presentation",
        "financial report",
        "MIS automation",
        "calculation accuracy",
        "data analysis",
        "advanced excel",
        "business calculation",
        "decimal control",
        "numeric formatting",
        "accounting formula"
    ],

    aliases:[
        "decimal remover",
        "number formatter",
        "rounding formula",
        "amount round function",
        "precision control"
    ],

    questions:[
        "How to round numbers?",
        "How to remove decimals?",
        "How to round currency values?",
        "How to use ROUND function?",
        "How to round percentage?",
        "How to format financial numbers?"
    ],

    dataset:"finance",

    category:"Math",

    formula:"=ROUND(A2,2)",

    syntax:"ROUND(number,num_digits)",

    explanation:"Rounds a number to a specified number of digits.",

    breakdown:[
        "A2 = Number",
        "2 = Decimal Places",
        "Returns Rounded Value"
    ],

    output:"1250.50",

    business:"Used in finance reports, invoices, and financial calculations.",

    tips:[
        "Use ROUNDUP for always upward rounding.",
        "Use ROUNDDOWN for always downward rounding."
    ],

    errors:[
        "#VALUE!",
        "Invalid number"
    ],

    alternative:"=ROUNDUP(A2,2)"
},

// ==========================================================
// ROUNDUP
// ==========================================================
{
    title:"ROUNDUP",

    keywords:[
        "roundup","always round up",
        "increase decimal","ceiling rounding",
        "financial rounding","amount rounding",
        "invoice calculation","billing formula",
        "currency rounding","quantity calculation",
        "number formatting","decimal control",
        "excel math function","finance report",
        "MIS calculation","pricing formula",
        "business calculation","advanced excel",
        "data preparation","numeric adjustment",
        "calculation accuracy","automation",
        "accounting formula","report formatting"
    ],

    aliases:[
        "always higher rounding",
        "upward rounding formula",
        "decimal increase formula",
        "ceiling number formula",
        "financial round up"
    ],

    questions:[
        "How to always round up?",
        "How to round amount upward?",
        "How to use ROUNDUP function?",
        "How to remove decimals by rounding?",
        "How to round invoice values?",
        "How to increase decimal rounding?"
    ],

    dataset:"finance",

    category:"Math",

    formula:"=ROUNDUP(A2,0)",

    syntax:"ROUNDUP(number,num_digits)",

    explanation:"Rounds a number upward away from zero.",

    breakdown:[
        "A2 = Number",
        "0 = No Decimal",
        "Always rounds upward"
    ],

    output:"126",

    business:"Used in billing, pricing, and financial calculations where values must always increase.",

    tips:[
        "Useful for quantity calculations.",
        "ROUNDUP never decreases the value."
    ],

    errors:[
        "#VALUE!",
        "Invalid number"
    ],

    alternative:"=CEILING(A2,1)"
},


// ==========================================================
// ROUNDDOWN
// ==========================================================
{
    title:"ROUNDDOWN",

    keywords:[
        "rounddown","always round down",
        "decrease decimal","floor rounding",
        "financial calculation","amount adjustment",
        "quantity rounding","inventory formula",
        "number formatting","decimal removal",
        "excel math function","finance report",
        "MIS calculation","business calculation",
        "pricing adjustment","data preparation",
        "numeric control","automation",
        "advanced excel","accounting formula",
        "calculation accuracy","report formatting"
    ],

    aliases:[
        "lower rounding formula",
        "downward rounding",
        "floor number formula",
        "decimal reduction formula",
        "always lower rounding"
    ],

    questions:[
        "How to always round down?",
        "How to remove decimals downward?",
        "How to use ROUNDDOWN function?",
        "How to reduce calculated values?",
        "How to round inventory quantity?",
        "How to control decimals?"
    ],

    dataset:"inventory",

    category:"Math",

    formula:"=ROUNDDOWN(A2,0)",

    syntax:"ROUNDDOWN(number,num_digits)",

    explanation:"Rounds a number downward toward zero.",

    breakdown:[
        "A2 = Number",
        "0 = Decimal Places",
        "Always decreases value"
    ],

    output:"125",

    business:"Used for inventory, quantity planning, and controlled financial calculations.",

    tips:[
        "Useful when values cannot exceed limits.",
        "ROUNDDOWN removes excess decimals."
    ],

    errors:[
        "#VALUE!",
        "Invalid number"
    ],

    alternative:"=FLOOR(A2,1)"
},


// ==========================================================
// MROUND
// ==========================================================
{
    title:"MROUND",

    keywords:[
        "mround","multiple rounding",
        "round to nearest multiple",
        "nearest 5","nearest 10",
        "pricing rounding","quantity rounding",
        "pack size calculation",
        "invoice adjustment",
        "financial rounding",
        "excel math function",
        "number formatting",
        "business calculation",
        "sales pricing",
        "inventory planning",
        "production quantity",
        "MIS report",
        "advanced excel",
        "automation",
        "numeric analysis",
        "decimal control"
    ],

    aliases:[
        "nearest multiple formula",
        "round to 5 formula",
        "round to 10 formula",
        "multiple calculator",
        "pack size rounding"
    ],

    questions:[
        "How to round to nearest 5?",
        "How to round to nearest 10?",
        "How to use MROUND function?",
        "How to calculate pack quantity?",
        "How to round prices?"
    ],

    dataset:"sales",

    category:"Math",

    formula:"=MROUND(A2,5)",

    syntax:"MROUND(number,multiple)",

    explanation:"Rounds a number to the nearest specified multiple.",

    breakdown:[
        "A2 = Number",
        "5 = Required Multiple",
        "Returns nearest multiple"
    ],

    output:"125",

    business:"Used for pricing, packaging, and standardized business calculations.",

    tips:[
        "Useful for rounding prices.",
        "Multiple can be 5,10,100 etc."
    ],

    errors:[
        "#NUM!",
        "Different sign values"
    ],

    alternative:"=ROUND(A2/5,0)*5"
},


// ==========================================================
// CEILING
// ==========================================================
{
    title:"CEILING",

    keywords:[
        "ceiling","round up multiple",
        "next multiple","upper rounding",
        "quantity calculation",
        "inventory planning",
        "pack size",
        "production planning",
        "financial rounding",
        "pricing formula",
        "excel math function",
        "number adjustment",
        "business calculation",
        "MIS reporting",
        "automation",
        "advanced excel",
        "numeric analysis",
        "capacity planning",
        "resource planning",
        "decimal control",
        "standardization"
    ],

    aliases:[
        "next multiple formula",
        "upper limit rounding",
        "minimum required quantity",
        "higher multiple rounding",
        "capacity rounding"
    ],

    questions:[
        "How to round up to nearest multiple?",
        "How to calculate pack quantity?",
        "How to use CEILING function?",
        "How to increase quantity automatically?",
        "How to round prices upward?"
    ],

    dataset:"inventory",

    category:"Math",

    formula:"=CEILING(A2,10)",

    syntax:"CEILING(number,significance)",

    explanation:"Rounds a number up to the nearest specified multiple.",

    breakdown:[
        "A2 = Number",
        "10 = Multiple",
        "Returns next multiple"
    ],

    output:"130",

    business:"Used in inventory, production planning, and quantity management.",

    tips:[
        "Useful when minimum quantity is required.",
        "Similar to ROUNDUP with multiples."
    ],

    errors:[
        "#NUM!",
        "Invalid significance"
    ],

    alternative:"=ROUNDUP(A2/10,0)*10"
},


// ==========================================================
// FLOOR
// ==========================================================
{
    title:"FLOOR",

    keywords:[
        "floor","round down multiple",
        "lower multiple","minimum value",
        "quantity rounding",
        "inventory control",
        "pricing calculation",
        "production planning",
        "financial rounding",
        "excel math function",
        "number adjustment",
        "business calculation",
        "MIS reporting",
        "automation",
        "advanced excel",
        "numeric analysis",
        "capacity calculation",
        "standard quantity",
        "decimal control",
        "data preparation"
    ],

    aliases:[
        "lower multiple formula",
        "previous multiple rounding",
        "downward multiple",
        "minimum quantity formula",
        "capacity rounding"
    ],

    questions:[
        "How to round down to nearest multiple?",
        "How to use FLOOR function?",
        "How to calculate lower quantity?",
        "How to control inventory quantity?",
        "How to round price downward?"
    ],

    dataset:"inventory",

    category:"Math",

    formula:"=FLOOR(A2,10)",

    syntax:"FLOOR(number,significance)",

    explanation:"Rounds a number down to the nearest specified multiple.",

    breakdown:[
        "A2 = Number",
        "10 = Multiple",
        "Returns lower multiple"
    ],

    output:"120",

    business:"Used for inventory limits, pricing rules, and controlled rounding.",

    tips:[
        "Useful when values should not exceed limits.",
        "Works opposite to CEILING."
    ],

    errors:[
        "#NUM!",
        "Invalid significance"
    ],

    alternative:"=ROUNDDOWN(A2/10,0)*10"
},

// ==========================================================
// IF
// ==========================================================
{
    title:"IF",

    keywords:[
        "if","condition","conditional formula",
        "check condition","decision formula",
        "yes no formula","pass fail",
        "approval status","salary check",
        "greater than","less than",
        "logical test","excel decision",
        "business rule","automation",
        "risk classification","employee eligibility",
        "status formula","mis reporting",
        "dashboard logic","advanced excel",
        "validation formula","comparison formula",
        "logical function","data analysis"
    ],

    aliases:[
        "condition formula",
        "decision formula",
        "if else formula",
        "logical test formula",
        "true false formula"
    ],

    questions:[
        "How to use IF formula?",
        "How to create yes no condition?",
        "How to check salary condition?",
        "How to create pass fail formula?",
        "How to assign status?",
        "How to apply business rules?",
        "How to compare values?"
    ],

    dataset:"employees",

    category:"Logical",

    formula:'=IF(C2>=50000,"High","Low")',

    syntax:"IF(logical_test,value_if_true,value_if_false)",

    explanation:"Checks a condition and returns one value if TRUE and another value if FALSE.",

    breakdown:[
        "C2 = Salary Value",
        ">=50000 = Condition",
        "High = TRUE Result",
        "Low = FALSE Result"
    ],

    output:"High",

    business:"Used for employee classification, approval logic, risk checks, and MIS automation.",

    tips:[
        "Can be nested for multiple decisions.",
        "Combine with AND/OR for complex logic."
    ],

    errors:[
        "#VALUE!",
        "Incorrect logical test"
    ],

    alternative:'=IFS(C2>=50000,"High",C2<50000,"Low")'
},


// ==========================================================
// IFS
// ==========================================================
{
    title:"IFS",

    keywords:[
        "ifs","multiple if condition",
        "multiple criteria",
        "nested if replacement",
        "grading formula",
        "salary slab",
        "performance rating",
        "risk category",
        "classification formula",
        "multiple decision",
        "logical function",
        "business rules",
        "approval matrix",
        "employee grading",
        "score classification",
        "dashboard logic",
        "advanced excel",
        "automation",
        "data analysis",
        "status mapping",
        "condition ranking",
        "decision tree"
    ],

    aliases:[
        "multiple condition if",
        "nested if alternative",
        "grading formula",
        "classification formula",
        "multi decision formula"
    ],

    questions:[
        "How to use multiple IF conditions?",
        "How to create grading formula?",
        "How to classify employees?",
        "How to replace nested IF?",
        "How to create salary slabs?",
        "How to assign performance rating?"
    ],

    dataset:"performance",

    category:"Logical",

    formula:'=IFS(B2>=90,"Excellent",B2>=75,"Good",B2>=50,"Average",B2<50,"Poor")',

    syntax:"IFS(condition1,result1,[condition2,result2],...)",

    explanation:"Tests multiple conditions and returns the first TRUE result.",

    breakdown:[
        "B2 = Performance Score",
        "90+ = Excellent",
        "75+ = Good",
        "50+ = Average"
    ],

    output:"Good",

    business:"Used for grading, employee ratings, risk classification, and approval matrices.",

    tips:[
        "Cleaner alternative to multiple nested IF.",
        "Conditions are checked from left to right."
    ],

    errors:[
        "#N/A",
        "No TRUE condition found"
    ],

    alternative:'=IF(B2>=90,"Excellent",IF(B2>=75,"Good","Average"))'
},


// ==========================================================
// AND
// ==========================================================
{
    title:"AND",

    keywords:[
        "and","multiple conditions",
        "all conditions true",
        "logical test",
        "combined condition",
        "eligibility check",
        "approval condition",
        "employee criteria",
        "loan eligibility",
        "risk assessment",
        "validation formula",
        "business rule",
        "logical function",
        "advanced excel",
        "automation",
        "decision making",
        "data validation",
        "compliance check",
        "multiple checks",
        "status formula",
        "MIS logic"
    ],

    aliases:[
        "all condition formula",
        "multiple true check",
        "combined logical test",
        "criteria validation",
        "joint condition formula"
    ],

    questions:[
        "How to check multiple conditions?",
        "How to verify all criteria?",
        "How to create approval formula?",
        "How to check employee eligibility?",
        "How to use AND function?",
        "How to combine conditions?"
    ],

    dataset:"employees",

    category:"Logical",

    formula:"=AND(C2>=50000,D2=\"Active\")",

    syntax:"AND(logical1,[logical2],...)",

    explanation:"Returns TRUE only when all conditions are TRUE.",

    breakdown:[
        "C2>=50000 = Salary Check",
        "D2=\"Active\" = Status Check",
        "Both must be TRUE"
    ],

    output:"TRUE",

    business:"Used for approval workflows, eligibility checks, and compliance validation.",

    tips:[
        "Usually combined with IF.",
        "All conditions must pass."
    ],

    errors:[
        "#VALUE!",
        "Invalid condition"
    ],

    alternative:'=IF(AND(C2>=50000,D2="Active"),"Eligible","Not Eligible")'
},


// ==========================================================
// OR
// ==========================================================
{
    title:"OR",

    keywords:[
        "or","any condition true",
        "multiple options",
        "alternative condition",
        "logical test",
        "eligibility formula",
        "exception check",
        "risk condition",
        "approval logic",
        "employee criteria",
        "validation formula",
        "business rules",
        "logical function",
        "advanced excel",
        "automation",
        "decision formula",
        "multiple choices",
        "status checking",
        "MIS logic",
        "data analysis"
    ],

    aliases:[
        "any condition formula",
        "alternative check",
        "multiple option logic",
        "one condition true",
        "choice formula"
    ],

    questions:[
        "How to check any condition?",
        "How to use OR formula?",
        "How to verify multiple options?",
        "How to create exception logic?",
        "How to check employee status?",
        "How to combine alternatives?"
    ],

    dataset:"employees",

    category:"Logical",

    formula:'=OR(D2="Active",D2="Pending")',

    syntax:"OR(logical1,[logical2],...)",

    explanation:"Returns TRUE if at least one condition is TRUE.",

    breakdown:[
        "D2=\"Active\" = First Condition",
        "D2=\"Pending\" = Second Condition",
        "Any TRUE gives TRUE"
    ],

    output:"TRUE",

    business:"Used for exception handling, status checks, and flexible approval rules.",

    tips:[
        "Combine with IF for output messages.",
        "Only one condition needs to match."
    ],

    errors:[
        "#VALUE!",
        "Invalid condition"
    ],

    alternative:'=IF(OR(D2="Active",D2="Pending"),"Open","Closed")'
},


// ==========================================================
// NOT
// ==========================================================
{
    title:"NOT",

    keywords:[
        "not","reverse condition",
        "opposite logic",
        "negative condition",
        "exclude criteria",
        "not equal check",
        "exception formula",
        "validation formula",
        "logical function",
        "error checking",
        "data quality",
        "compliance check",
        "approval logic",
        "business rules",
        "advanced excel",
        "automation",
        "filter condition",
        "status validation",
        "logical reversal",
        "MIS reporting"
    ],

    aliases:[
        "reverse logic formula",
        "opposite condition",
        "negative check",
        "exclude formula",
        "inverse logic"
    ],

    questions:[
        "How to reverse condition?",
        "How to check not equal?",
        "How to exclude values?",
        "How to use NOT function?",
        "How to create exception check?",
        "How to reverse TRUE FALSE?"
    ],

    dataset:"transactions",

    category:"Logical",

    formula:'=NOT(B2="Closed")',

    syntax:"NOT(logical)",

    explanation:"Reverses the logical value. TRUE becomes FALSE and FALSE becomes TRUE.",

    breakdown:[
        "B2=\"Closed\" = Original Condition",
        "NOT reverses result",
        "Returns opposite value"
    ],

    output:"TRUE",

    business:"Used for exception reporting, validation, and compliance checks.",

    tips:[
        "Useful with IF and AND.",
        "Commonly used for excluding conditions."
    ],

    errors:[
        "#VALUE!",
        "Invalid logical value"
    ],

    alternative:'=B2<>"Closed"'
},

// ==========================================================
// IFERROR
// ==========================================================
{
    title:"IFERROR",

    keywords:[
        "iferror","remove error",
        "handle error","replace error",
        "hide formula error",
        "#n/a fix",
        "#value fix",
        "#div/0 fix",
        "error handling",
        "blank instead error",
        "excel error formula",
        "lookup error",
        "formula validation",
        "clean report",
        "dashboard formula",
        "mis reporting",
        "advanced excel",
        "data quality",
        "automation",
        "exception handling",
        "formula protection",
        "business reporting",
        "error management"
    ],

    aliases:[
        "error remover",
        "error handling formula",
        "formula error fix",
        "hide errors",
        "safe formula"
    ],

    questions:[
        "How to remove formula errors?",
        "How to replace error with blank?",
        "How to handle #N/A error?",
        "How to use IFERROR?",
        "How to fix lookup errors?",
        "How to show custom error message?"
    ],

    dataset:"lookup",

    category:"Logical",

    formula:'=IFERROR(VLOOKUP(H2,A:D,4,FALSE),"Not Found")',

    syntax:"IFERROR(value,value_if_error)",

    explanation:"Returns a custom result when a formula generates an error.",

    breakdown:[
        "VLOOKUP = Main Formula",
        "\"Not Found\" = Error Result",
        "Prevents error display"
    ],

    output:"Not Found",

    business:"Used in dashboards, reports, and lookup formulas to create clean outputs.",

    tips:[
        "Useful with VLOOKUP, XLOOKUP, and calculations.",
        "Avoid showing technical errors to users."
    ],

    errors:[
        "#VALUE!",
        "Wrong formula logic"
    ],

    alternative:'=IF(ISERROR(A2),"Error",A2)'
},


// ==========================================================
// IFNA
// ==========================================================
{
    title:"IFNA",

    keywords:[
        "ifna","handle na error",
        "#n/a fix",
        "lookup missing value",
        "replace not found",
        "vlookup error",
        "xlookup error",
        "missing data handling",
        "error handling",
        "lookup validation",
        "clean output",
        "excel logical function",
        "data quality",
        "report automation",
        "advanced excel",
        "dashboard formula",
        "exception handling",
        "missing record",
        "master data check",
        "formula protection"
    ],

    aliases:[
        "n/a error handler",
        "not found handler",
        "lookup missing value fix",
        "missing data formula"
    ],

    questions:[
        "How to fix #N/A error?",
        "How to replace not found?",
        "How to handle lookup missing values?",
        "How to use IFNA?",
        "How to show custom message for missing data?"
    ],

    dataset:"employees",

    category:"Logical",

    formula:'=IFNA(XLOOKUP(H2,A:A,D:D),"Employee Not Found")',

    syntax:"IFNA(value,value_if_na)",

    explanation:"Handles only #N/A errors and returns a custom result.",

    breakdown:[
        "XLOOKUP = Lookup Formula",
        "#N/A = Missing Match",
        "Custom Message Returned"
    ],

    output:"Employee Not Found",

    business:"Used for employee lookup, master data validation, and clean reporting.",

    tips:[
        "Use IFNA when you only want to handle missing values.",
        "Use IFERROR for all types of errors."
    ],

    errors:[
        "#VALUE!",
        "Invalid lookup"
    ],

    alternative:'=IFERROR(XLOOKUP(H2,A:A,D:D),"Not Found")'
},


// ==========================================================
// TRUE
// ==========================================================
{
    title:"TRUE",

    keywords:[
        "true","logical true",
        "true value",
        "boolean value",
        "condition result",
        "logical test",
        "validation formula",
        "yes value",
        "approval flag",
        "excel logical function",
        "automation",
        "status indicator",
        "data validation",
        "formula logic",
        "dashboard status",
        "decision making",
        "advanced excel",
        "binary value",
        "comparison result",
        "business rules"
    ],

    aliases:[
        "true function",
        "boolean true",
        "positive logic value",
        "yes indicator"
    ],

    questions:[
        "How to return TRUE?",
        "How to use TRUE function?",
        "What is boolean TRUE?",
        "How to create logical flag?",
        "How to use TRUE in formulas?"
    ],

    dataset:"validation",

    category:"Logical",

    formula:"=TRUE()",

    syntax:"TRUE()",

    explanation:"Returns the logical value TRUE.",

    breakdown:[
        "No arguments required",
        "Returns TRUE value"
    ],

    output:"TRUE",

    business:"Used for logical testing, flags, and automation rules.",

    tips:[
        "TRUE equals 1 in some calculations.",
        "Often used with logical formulas."
    ],

    errors:[
        "No common errors"
    ],

    alternative:"=1=1"
},


// ==========================================================
// FALSE
// ==========================================================
{
    title:"FALSE",

    keywords:[
        "false","logical false",
        "false value",
        "boolean false",
        "condition failure",
        "negative result",
        "validation formula",
        "approval rejection",
        "logical test",
        "excel logical function",
        "automation",
        "status indicator",
        "formula logic",
        "dashboard status",
        "decision making",
        "advanced excel",
        "binary value",
        "comparison result",
        "business rules",
        "data validation"
    ],

    aliases:[
        "false function",
        "boolean false",
        "negative logic value",
        "no indicator"
    ],

    questions:[
        "How to return FALSE?",
        "How to use FALSE function?",
        "What is boolean FALSE?",
        "How to create rejection flag?",
        "How to use FALSE in formulas?"
    ],

    dataset:"validation",

    category:"Logical",

    formula:"=FALSE()",

    syntax:"FALSE()",

    explanation:"Returns the logical value FALSE.",

    breakdown:[
        "No arguments required",
        "Returns FALSE value"
    ],

    output:"FALSE",

    business:"Used for logical checks, rejection flags, and validation rules.",

    tips:[
        "FALSE equals 0 in calculations.",
        "Commonly used with IF formulas."
    ],

    errors:[
        "No common errors"
    ],

    alternative:"=1=2"
},


// ==========================================================
// SWITCH
// ==========================================================
{
    title:"SWITCH",

    keywords:[
        "switch","multiple value mapping",
        "case formula",
        "replace nested if",
        "status mapping",
        "category mapping",
        "employee grade",
        "risk category",
        "approval status",
        "multiple choices",
        "logical function",
        "classification formula",
        "dashboard mapping",
        "business rules",
        "advanced excel",
        "automation",
        "data transformation",
        "decision table",
        "report logic",
        "status conversion",
        "condition mapping"
    ],

    aliases:[
        "case statement excel",
        "multiple choice formula",
        "mapping formula",
        "nested if alternative",
        "value converter"
    ],

    questions:[
        "How to replace multiple IF?",
        "How to map values?",
        "How to create status mapping?",
        "How to use SWITCH function?",
        "How to assign categories?",
        "How to create grade formula?"
    ],

    dataset:"employees",

    category:"Logical",

    formula:'=SWITCH(B2,"A","Excellent","B","Good","C","Average","Poor")',

    syntax:"SWITCH(expression,value1,result1,[default])",

    explanation:"Matches a value against multiple options and returns the corresponding result.",

    breakdown:[
        "B2 = Grade Value",
        "A/B/C = Matching Values",
        "Returns Category"
    ],

    output:"Excellent",

    business:"Used for grade mapping, status conversion, and replacing long nested IF formulas.",

    tips:[
        "Cleaner than multiple IF conditions.",
        "Use default value for unmatched cases."
    ],

    errors:[
        "#N/A",
        "No matching value"
    ],

    alternative:'=IFS(B2="A","Excellent",B2="B","Good")'
},

// ==========================================================
// LEFT
// ==========================================================
{
    title:"LEFT",

    keywords:[
        "left","extract left characters",
        "get first characters",
        "extract starting text",
        "employee code extraction",
        "prefix extraction",
        "text split",
        "remove right side text",
        "first letters",
        "product code",
        "excel text function",
        "data cleaning",
        "text manipulation",
        "string extraction",
        "automation",
        "advanced excel",
        "report preparation",
        "text analysis",
        "id extraction",
        "data transformation",
        "import cleaning",
        "master data cleaning"
    ],

    aliases:[
        "first character formula",
        "starting text extractor",
        "prefix finder",
        "left side extraction",
        "beginning text formula"
    ],

    questions:[
        "How to extract first characters?",
        "How to get employee code?",
        "How to extract text from left?",
        "How to remove ending text?",
        "How to use LEFT function?",
        "How to split codes from text?"
    ],

    dataset:"employees",

    category:"Text",

    formula:"=LEFT(A2,3)",

    syntax:"LEFT(text,[num_chars])",

    explanation:"Returns the specified number of characters from the beginning of a text string.",

    breakdown:[
        "A2 = Source Text",
        "3 = Number of Characters",
        "Returns first 3 characters"
    ],

    output:"EMP",

    business:"Used for employee IDs, product codes, and data extraction.",

    tips:[
        "Useful when codes have fixed length.",
        "Combine with LEN for dynamic extraction."
    ],

    errors:[
        "#VALUE!",
        "Invalid character count"
    ],

    alternative:"=TEXTBEFORE(A2,\"-\")"
},


// ==========================================================
// RIGHT
// ==========================================================
{
    title:"RIGHT",

    keywords:[
        "right","extract right characters",
        "get ending characters",
        "last characters",
        "suffix extraction",
        "account number last digits",
        "employee code",
        "text extraction",
        "remove left text",
        "excel text function",
        "data cleaning",
        "string manipulation",
        "automation",
        "advanced excel",
        "report preparation",
        "id extraction",
        "text analysis",
        "data transformation",
        "master data cleaning",
        "imported data cleaning"
    ],

    aliases:[
        "last character formula",
        "ending text extractor",
        "suffix finder",
        "right side extraction",
        "end text formula"
    ],

    questions:[
        "How to extract last characters?",
        "How to get last digits?",
        "How to extract text from right?",
        "How to use RIGHT function?",
        "How to split account numbers?",
        "How to extract suffix?"
    ],

    dataset:"customers",

    category:"Text",

    formula:"=RIGHT(A2,4)",

    syntax:"RIGHT(text,[num_chars])",

    explanation:"Returns characters from the end of a text string.",

    breakdown:[
        "A2 = Source Text",
        "4 = Characters Required",
        "Returns ending characters"
    ],

    output:"4521",

    business:"Used for account numbers, IDs, and suffix extraction.",

    tips:[
        "Useful for extracting last digits.",
        "Works well with LEFT and MID."
    ],

    errors:[
        "#VALUE!",
        "Invalid character count"
    ],

    alternative:"=TEXTAFTER(A2,\"-\")"
},


// ==========================================================
// MID
// ==========================================================
{
    title:"MID",

    keywords:[
        "mid","extract middle text",
        "text from position",
        "extract characters",
        "employee code",
        "account number extraction",
        "substring formula",
        "string extraction",
        "text parsing",
        "excel text function",
        "data cleaning",
        "automation",
        "advanced excel",
        "text analysis",
        "data transformation",
        "import processing",
        "master data",
        "report preparation",
        "dynamic extraction",
        "character extraction"
    ],

    aliases:[
        "substring formula",
        "middle character extractor",
        "text position formula",
        "character extraction",
        "partial text formula"
    ],

    questions:[
        "How to extract middle characters?",
        "How to extract text from position?",
        "How to use MID function?",
        "How to split text?",
        "How to extract employee code?",
        "How to get substring?"
    ],

    dataset:"employees",

    category:"Text",

    formula:"=MID(A2,4,5)",

    syntax:"MID(text,start_num,num_chars)",

    explanation:"Returns characters from the middle of a text string based on starting position.",

    breakdown:[
        "A2 = Source Text",
        "4 = Starting Position",
        "5 = Characters Required"
    ],

    output:"12345",

    business:"Used for extracting parts of IDs, codes, and structured text.",

    tips:[
        "Very useful for fixed-format data.",
        "Combine with FIND for dynamic extraction."
    ],

    errors:[
        "#VALUE!",
        "Invalid start position"
    ],

    alternative:'=TEXTBEFORE(TEXTAFTER(A2,"-"),"-")'
},


// ==========================================================
// LEN
// ==========================================================
{
    title:"LEN",

    keywords:[
        "len","length of text",
        "count characters",
        "text length",
        "character count",
        "data validation",
        "check id length",
        "password length",
        "employee code length",
        "text analysis",
        "excel text function",
        "data quality",
        "cleaning data",
        "string calculation",
        "automation",
        "advanced excel",
        "validation formula",
        "master data check",
        "text processing",
        "report preparation"
    ],

    aliases:[
        "character counter",
        "text size formula",
        "string length",
        "length checker",
        "text validator"
    ],

    questions:[
        "How to count characters?",
        "How to check text length?",
        "How to validate employee ID?",
        "How to use LEN function?",
        "How many characters exist?",
        "How to find blank spaces?"
    ],

    dataset:"validation",

    category:"Text",

    formula:"=LEN(A2)",

    syntax:"LEN(text)",

    explanation:"Returns the number of characters in a text string including spaces.",

    breakdown:[
        "A2 = Text Value",
        "Counts every character",
        "Returns length"
    ],

    output:"10",

    business:"Used for data validation, ID checks, and cleaning imported files.",

    tips:[
        "Combine with TRIM to remove unwanted spaces.",
        "Useful in data quality checks."
    ],

    errors:[
        "#VALUE!"
    ],

    alternative:"=COUNTA(A2)"
},


// ==========================================================
// TRIM
// ==========================================================
{
    title:"TRIM",

    keywords:[
        "trim","remove spaces",
        "extra spaces removal",
        "clean text",
        "data cleaning",
        "remove unwanted spaces",
        "duplicate matching",
        "lookup error fix",
        "employee data cleaning",
        "imported data cleanup",
        "text standardization",
        "excel text function",
        "master data cleaning",
        "automation",
        "advanced excel",
        "data quality",
        "string cleaning",
        "report preparation",
        "text formatting",
        "database cleanup"
    ],

    aliases:[
        "space remover",
        "text cleaner",
        "remove extra spaces",
        "clean string formula",
        "data cleanup formula"
    ],

    questions:[
        "How to remove extra spaces?",
        "How to clean imported data?",
        "How to fix lookup mismatch?",
        "How to use TRIM function?",
        "How to standardize text?",
        "How to remove unwanted spaces?"
    ],

    dataset:"customer_master",

    category:"Text",

    formula:"=TRIM(A2)",

    syntax:"TRIM(text)",

    explanation:"Removes extra spaces from text except single spaces between words.",

    breakdown:[
        "A2 = Original Text",
        "Removes leading spaces",
        "Removes extra spaces"
    ],

    output:"Sunil Kumar",

    business:"Used for cleaning master data and fixing lookup matching issues.",

    tips:[
        "Very useful before VLOOKUP/XLOOKUP.",
        "Does not remove non-printing characters."
    ],

    errors:[
        "#VALUE!"
    ],

    alternative:"=CLEAN(TRIM(A2))"
},

// ==========================================================
// CONCAT
// ==========================================================
{
    title:"CONCAT",

    keywords:[
        "concat","combine text",
        "join text","merge cells",
        "combine columns",
        "create full name",
        "combine employee details",
        "text joining",
        "string combination",
        "excel text function",
        "data preparation",
        "email creation",
        "report formatting",
        "automation",
        "advanced excel",
        "text manipulation",
        "master data",
        "data transformation",
        "dynamic text",
        "text analysis",
        "customer details",
        "employee information"
    ],

    aliases:[
        "text combine formula",
        "merge text formula",
        "join columns",
        "string concatenate",
        "combine values"
    ],

    questions:[
        "How to combine text?",
        "How to merge first and last name?",
        "How to join columns?",
        "How to use CONCAT function?",
        "How to create full name?",
        "How to combine employee details?"
    ],

    dataset:"employees",

    category:"Text",

    formula:"=CONCAT(A2,B2)",

    syntax:"CONCAT(text1,[text2],...)",

    explanation:"Combines multiple text values into one text string.",

    breakdown:[
        "A2 = First Name",
        "B2 = Last Name",
        "Combines both values"
    ],

    output:"SunilKumar",

    business:"Used for creating names, IDs, and combined business fields.",

    tips:[
        "Use TEXTJOIN when separators are required.",
        "Modern replacement for CONCATENATE."
    ],

    errors:[
        "#VALUE!",
        "Invalid reference"
    ],

    alternative:'=A2&" "&B2'
},


// ==========================================================
// CONCATENATE
// ==========================================================
{
    title:"CONCATENATE",

    keywords:[
        "concatenate","combine text",
        "join strings",
        "merge values",
        "combine cells",
        "create full name",
        "append text",
        "text joining",
        "old concatenate function",
        "excel text formula",
        "employee data",
        "customer data",
        "email generation",
        "report creation",
        "automation",
        "advanced excel",
        "data preparation",
        "string manipulation",
        "text formatting",
        "dynamic messages"
    ],

    aliases:[
        "join text formula",
        "combine strings",
        "merge cell values",
        "append formula"
    ],

    questions:[
        "How to combine two cells?",
        "How to join text with formula?",
        "How to use CONCATENATE?",
        "How to create messages?",
        "How to merge employee details?"
    ],

    dataset:"employees",

    category:"Text",

    formula:'=CONCATENATE(A2," ",B2)',

    syntax:"CONCATENATE(text1,text2,...)",

    explanation:"Joins multiple text strings into one text value.",

    breakdown:[
        "A2 = First Name",
        "\" \" = Space",
        "B2 = Last Name"
    ],

    output:"Sunil Kumar",

    business:"Used for creating full names, descriptions, and combined fields.",

    tips:[
        "CONCAT is the newer replacement.",
        "Use & operator for simple joining."
    ],

    errors:[
        "#VALUE!"
    ],

    alternative:'=A2&" "&B2'
},


// ==========================================================
// TEXTJOIN
// ==========================================================
{
    title:"TEXTJOIN",

    keywords:[
        "textjoin","combine text with separator",
        "merge multiple cells",
        "join values",
        "comma separated list",
        "combine employee data",
        "dynamic text list",
        "ignore blanks",
        "text aggregation",
        "excel text function",
        "report preparation",
        "automation",
        "advanced excel",
        "data transformation",
        "string joining",
        "dashboard text",
        "email list creation",
        "customer list",
        "text formatting",
        "business reporting"
    ],

    aliases:[
        "separator join formula",
        "comma combine formula",
        "multiple cell merge",
        "dynamic concatenate",
        "text merger"
    ],

    questions:[
        "How to combine multiple cells?",
        "How to join values with comma?",
        "How to ignore blanks while joining?",
        "How to use TEXTJOIN?",
        "How to create comma separated list?"
    ],

    dataset:"customers",

    category:"Text",

    formula:'=TEXTJOIN(", ",TRUE,A2:A10)',

    syntax:"TEXTJOIN(delimiter,ignore_empty,text1,...)",

    explanation:"Combines multiple text values using a specified separator.",

    breakdown:[
        "\", \" = Separator",
        "TRUE = Ignore blanks",
        "A2:A10 = Text Range"
    ],

    output:"India, USA, UK",

    business:"Used for creating lists, reports, email groups, and summaries.",

    tips:[
        "Best option for joining multiple cells.",
        "Works with dynamic ranges."
    ],

    errors:[
        "#VALUE!"
    ],

    alternative:'=CONCAT(A2:A10)'
},


// ==========================================================
// UPPER
// ==========================================================
{
    title:"UPPER",

    keywords:[
        "upper","uppercase",
        "convert capital letters",
        "text formatting",
        "standardize text",
        "employee name formatting",
        "customer data cleaning",
        "master data cleanup",
        "text conversion",
        "excel text function",
        "data quality",
        "report formatting",
        "automation",
        "advanced excel",
        "string formatting",
        "database cleanup",
        "text normalization",
        "consistent naming",
        "data preparation"
    ],

    aliases:[
        "capital letter formula",
        "uppercase converter",
        "text capitalizer",
        "name formatter"
    ],

    questions:[
        "How to convert text to uppercase?",
        "How to make all letters capital?",
        "How to standardize names?",
        "How to use UPPER function?",
        "How to clean inconsistent text?"
    ],

    dataset:"employees",

    category:"Text",

    formula:"=UPPER(A2)",

    syntax:"UPPER(text)",

    explanation:"Converts all letters in a text string to uppercase.",

    breakdown:[
        "A2 = Original Text",
        "Converts lowercase letters",
        "Returns capital text"
    ],

    output:"SUNIL KUMAR",

    business:"Used for standardizing names, codes, and master data.",

    tips:[
        "Use PROPER for name format.",
        "Useful before comparison."
    ],

    errors:[
        "#VALUE!"
    ],

    alternative:"=PROPER(A2)"
},


// ==========================================================
// LOWER
// ==========================================================
{
    title:"LOWER",

    keywords:[
        "lower","lowercase",
        "convert small letters",
        "text formatting",
        "email formatting",
        "standardize text",
        "employee email",
        "customer data cleaning",
        "master data cleanup",
        "text conversion",
        "excel text function",
        "data quality",
        "automation",
        "advanced excel",
        "string formatting",
        "text normalization",
        "consistent data",
        "report preparation",
        "data transformation"
    ],

    aliases:[
        "small letter formula",
        "lowercase converter",
        "email formatter",
        "text standardizer"
    ],

    questions:[
        "How to convert text to lowercase?",
        "How to make email IDs lowercase?",
        "How to standardize text?",
        "How to use LOWER function?",
        "How to clean text format?"
    ],

    dataset:"employees",

    category:"Text",

    formula:"=LOWER(A2)",

    syntax:"LOWER(text)",

    explanation:"Converts all letters in a text string to lowercase.",

    breakdown:[
        "A2 = Original Text",
        "Converts capital letters",
        "Returns lowercase text"
    ],

    output:"sunil kumar",

    business:"Used for email formatting, text standardization, and data cleaning.",

    tips:[
        "Useful before duplicate checking.",
        "Combine with TRIM for clean data."
    ],

    errors:[
        "#VALUE!"
    ],

    alternative:"=PROPER(A2)"
},

// ==========================================================
// PROPER
// ==========================================================
{
    title:"PROPER",

    keywords:[
        "proper","capitalize text",
        "first letter capital",
        "name formatting",
        "employee name cleanup",
        "customer name format",
        "title case",
        "text standardization",
        "convert lowercase to proper",
        "excel text function",
        "data cleaning",
        "master data cleanup",
        "report formatting",
        "automation",
        "advanced excel",
        "string formatting",
        "customer database",
        "employee database",
        "text correction",
        "data quality",
        "name standardization"
    ],

    aliases:[
        "title case formula",
        "capitalize name formula",
        "name formatter",
        "proper case converter"
    ],

    questions:[
        "How to capitalize names?",
        "How to format employee names?",
        "How to convert text into proper case?",
        "How to use PROPER function?",
        "How to clean customer names?",
        "How to fix uppercase lowercase issue?"
    ],

    dataset:"employees",

    category:"Text",

    formula:"=PROPER(A2)",

    syntax:"PROPER(text)",

    explanation:"Converts the first letter of each word into uppercase and remaining letters into lowercase.",

    breakdown:[
        "A2 = Original Text",
        "First letter becomes capital",
        "Remaining letters become lowercase"
    ],

    output:"Sunil Kumar",

    business:"Used for employee master cleaning, customer databases, and reporting formats.",

    tips:[
        "Combine with TRIM for better cleaning.",
        "Useful after importing external data."
    ],

    errors:[
        "#VALUE!"
    ],

    alternative:"=UPPER(LEFT(A2,1))&LOWER(MID(A2,2,LEN(A2)))"
},


// ==========================================================
// FIND
// ==========================================================
{
    title:"FIND",

    keywords:[
        "find","locate text",
        "find character position",
        "search text position",
        "case sensitive search",
        "text location",
        "character finder",
        "extract text",
        "string position",
        "excel text function",
        "data parsing",
        "text analysis",
        "automation",
        "advanced excel",
        "data cleaning",
        "email extraction",
        "code extraction",
        "text validation",
        "string analysis",
        "position formula"
    ],

    aliases:[
        "character position formula",
        "text locator",
        "string finder",
        "position finder"
    ],

    questions:[
        "How to find text position?",
        "How to locate a character?",
        "How to find hyphen position?",
        "How to use FIND function?",
        "How to extract text before symbol?",
        "How to locate email @ position?"
    ],

    dataset:"customers",

    category:"Text",

    formula:'=FIND("@",A2)',

    syntax:"FIND(find_text,within_text,[start_num])",

    explanation:"Returns the position of a specific character or text inside another text.",

    breakdown:[
        "\"@\" = Character to Find",
        "A2 = Source Text",
        "Returns Position Number"
    ],

    output:"6",

    business:"Used for email parsing, code extraction, and text analysis.",

    tips:[
        "FIND is case-sensitive.",
        "Use SEARCH for non-case-sensitive search."
    ],

    errors:[
        "#VALUE!",
        "Text not found"
    ],

    alternative:'=SEARCH("@",A2)'
},


// ==========================================================
// SEARCH
// ==========================================================
{
    title:"SEARCH",

    keywords:[
        "search","find text",
        "locate word",
        "text search",
        "case insensitive search",
        "find keyword",
        "contains text",
        "string search",
        "text validation",
        "excel text function",
        "data analysis",
        "classification formula",
        "automation",
        "advanced excel",
        "customer feedback analysis",
        "comment analysis",
        "text filtering",
        "keyword detection",
        "data cleaning",
        "report preparation"
    ],

    aliases:[
        "keyword finder",
        "text contains formula",
        "case insensitive find",
        "string search formula"
    ],

    questions:[
        "How to check if text contains word?",
        "How to search text?",
        "How to find keyword in sentence?",
        "How to use SEARCH function?",
        "How to identify matching text?",
        "How to filter comments?"
    ],

    dataset:"feedback",

    category:"Text",

    formula:'=SEARCH("urgent",A2)',

    syntax:"SEARCH(find_text,within_text,[start_num])",

    explanation:"Finds the position of text within another text without considering uppercase or lowercase.",

    breakdown:[
        "\"urgent\" = Keyword",
        "A2 = Text Source",
        "Returns Position"
    ],

    output:"5",

    business:"Used for keyword detection, feedback analysis, and text classification.",

    tips:[
        "Not case-sensitive.",
        "Combine with ISNUMBER for TRUE/FALSE checks."
    ],

    errors:[
        "#VALUE!",
        "Text not found"
    ],

    alternative:'=ISNUMBER(SEARCH("urgent",A2))'
},


// ==========================================================
// SUBSTITUTE
// ==========================================================
{
    title:"SUBSTITUTE",

    keywords:[
        "substitute","replace text",
        "replace characters",
        "remove unwanted text",
        "change word",
        "data cleaning",
        "text correction",
        "replace symbol",
        "remove hyphen",
        "format correction",
        "excel text function",
        "master data cleanup",
        "automation",
        "advanced excel",
        "string replacement",
        "text transformation",
        "data preparation",
        "report formatting",
        "text standardization",
        "import cleanup"
    ],

    aliases:[
        "text replacement formula",
        "replace word formula",
        "character replacement",
        "string cleaner"
    ],

    questions:[
        "How to replace text?",
        "How to remove special characters?",
        "How to replace old value?",
        "How to use SUBSTITUTE function?",
        "How to clean imported data?",
        "How to replace symbols?"
    ],

    dataset:"customer_data",

    category:"Text",

    formula:'=SUBSTITUTE(A2,"-","")',

    syntax:"SUBSTITUTE(text,old_text,new_text,[instance_num])",

    explanation:"Replaces existing text with new text in a string.",

    breakdown:[
        "A2 = Original Text",
        "\"-\" = Text to Replace",
        "\"\" = Remove Character"
    ],

    output:"123456789",

    business:"Used for cleaning IDs, phone numbers, and imported files.",

    tips:[
        "Can replace specific occurrence using instance_num.",
        "Useful with data cleaning."
    ],

    errors:[
        "#VALUE!"
    ],

    alternative:'=REPLACE(A2,4,1,"")'
},


// ==========================================================
// REPLACE
// ==========================================================
{
    title:"REPLACE",

    keywords:[
        "replace","replace characters",
        "change text position",
        "overwrite text",
        "modify string",
        "mask data",
        "hide sensitive data",
        "change employee code",
        "text correction",
        "excel text function",
        "data cleaning",
        "automation",
        "advanced excel",
        "string manipulation",
        "text transformation",
        "report preparation",
        "data masking",
        "customer information",
        "text editing",
        "format correction"
    ],

    aliases:[
        "position replacement formula",
        "text overwrite formula",
        "string edit function",
        "character replacement"
    ],

    questions:[
        "How to replace characters by position?",
        "How to hide account digits?",
        "How to modify text?",
        "How to use REPLACE function?",
        "How to mask sensitive data?",
        "How to change part of text?"
    ],

    dataset:"accounts",

    category:"Text",

    formula:'=REPLACE(A2,1,4,"XXXX")',

    syntax:"REPLACE(old_text,start_num,num_chars,new_text)",

    explanation:"Replaces part of a text string based on position.",

    breakdown:[
        "A2 = Original Text",
        "1 = Starting Position",
        "4 = Characters Replace",
        "XXXX = New Text"
    ],

    output:"XXXX5678",

    business:"Used for masking account numbers, editing codes, and text corrections.",

    tips:[
        "Use SUBSTITUTE when replacing exact text.",
        "Use REPLACE when position is fixed."
    ],

    errors:[
        "#VALUE!",
        "Invalid position"
    ],

    alternative:'=SUBSTITUTE(A2,"1234","XXXX")'
},

// ==========================================================
// XLOOKUP
// ==========================================================
{
    title:"XLOOKUP",

    keywords:[
        "xlookup",
        "advanced lookup",
        "vlookup alternative",
        "find value",
        "return matching value",
        "dynamic lookup",
        "employee lookup",
        "salary lookup",
        "modern lookup",
        "excel lookup function",
        "two way lookup",
        "horizontal lookup",
        "vertical lookup",
        "exact match lookup",
        "data retrieval",
        "master data search",
        "automation",
        "advanced excel",
        "mis report",
        "dashboard lookup",
        "business analysis",
        "replacement for vlookup"
    ],

    aliases:[
        "vlookup replacement",
        "modern lookup formula",
        "dynamic vlookup",
        "smart lookup",
        "flexible lookup"
    ],

    questions:[
        "How to use XLOOKUP?",
        "What is replacement of VLOOKUP?",
        "How to find employee salary?",
        "How to lookup value from another table?",
        "How to create dynamic lookup?",
        "How to return matching data?"
    ],

    dataset:"employees",

    category:"Lookup",

    formula:"=XLOOKUP(H2,A:A,D:D)",

    syntax:"XLOOKUP(lookup_value,lookup_array,return_array)",

    explanation:"Searches a value in a range and returns the corresponding value from another range.",

    breakdown:[
        "H2 = Lookup Employee ID",
        "A:A = Employee ID Column",
        "D:D = Salary Column"
    ],

    output:"65000",

    business:"Used for employee details, salary reports, and dynamic dashboards.",

    tips:[
        "Can lookup left or right.",
        "Better than VLOOKUP for modern Excel."
    ],

    errors:[
        "#N/A",
        "Value not found"
    ],

    alternative:"=IFERROR(VLOOKUP(H2,A:D,4,FALSE),\"Not Found\")"
},


// ==========================================================
// INDEX
// ==========================================================
{
    title:"INDEX",

    keywords:[
        "index",
        "return value",
        "find position value",
        "lookup formula",
        "dynamic reference",
        "row column lookup",
        "advanced lookup",
        "table extraction",
        "data retrieval",
        "employee details",
        "salary extraction",
        "excel reference function",
        "two dimensional lookup",
        "dashboard formula",
        "automation",
        "advanced excel",
        "mis reporting",
        "database lookup",
        "dynamic formula",
        "business analysis"
    ],

    aliases:[
        "value extractor",
        "table lookup formula",
        "position based lookup",
        "dynamic return formula"
    ],

    questions:[
        "How to return value by position?",
        "How to use INDEX function?",
        "How to extract data from table?",
        "How to create dynamic lookup?",
        "How to get row column value?"
    ],

    dataset:"employees",

    category:"Lookup",

    formula:"=INDEX(D:D,5)",

    syntax:"INDEX(array,row_num,[column_num])",

    explanation:"Returns a value from a specified row and column position.",

    breakdown:[
        "D:D = Data Column",
        "5 = Row Number",
        "Returns matching value"
    ],

    output:"70000",

    business:"Used for dynamic reporting and advanced lookup solutions.",

    tips:[
        "Usually combined with MATCH.",
        "Supports two-dimensional lookup."
    ],

    errors:[
        "#REF!",
        "Invalid row number"
    ],

    alternative:"=XLOOKUP(A5,A:A,D:D)"
},


// ==========================================================
// MATCH
// ==========================================================
{
    title:"MATCH",

    keywords:[
        "match",
        "find position",
        "lookup position",
        "row number finder",
        "column position",
        "search value position",
        "index match",
        "dynamic lookup",
        "excel reference function",
        "data analysis",
        "employee search",
        "table matching",
        "automation",
        "advanced excel",
        "dashboard formula",
        "position finder",
        "exact match",
        "approximate match",
        "business reporting",
        "data retrieval"
    ],

    aliases:[
        "position lookup",
        "row finder",
        "column finder",
        "lookup index"
    ],

    questions:[
        "How to find position of value?",
        "How to use MATCH formula?",
        "How to find row number?",
        "How to create INDEX MATCH?",
        "How to search data position?"
    ],

    dataset:"employees",

    category:"Lookup",

    formula:"=MATCH(H2,A:A,0)",

    syntax:"MATCH(lookup_value,lookup_array,[match_type])",

    explanation:"Returns the relative position of a value in a range.",

    breakdown:[
        "H2 = Search Value",
        "A:A = Search Range",
        "0 = Exact Match"
    ],

    output:"5",

    business:"Used with INDEX for powerful lookup solutions.",

    tips:[
        "0 means exact match.",
        "Commonly paired with INDEX."
    ],

    errors:[
        "#N/A",
        "No matching value"
    ],

    alternative:"=XMATCH(H2,A:A)"
},


// ==========================================================
// XMATCH
// ==========================================================
{
    title:"XMATCH",

    keywords:[
        "xmatch",
        "advanced match",
        "modern match",
        "find position",
        "dynamic position lookup",
        "index xmatch",
        "lookup position",
        "exact search",
        "reverse lookup",
        "flexible matching",
        "excel lookup function",
        "data retrieval",
        "employee search",
        "table analysis",
        "automation",
        "advanced excel",
        "dashboard formula",
        "business reporting",
        "data matching",
        "reference function"
    ],

    aliases:[
        "match replacement",
        "modern match formula",
        "dynamic position finder",
        "smart match"
    ],

    questions:[
        "How to use XMATCH?",
        "What is replacement of MATCH?",
        "How to find position dynamically?",
        "How to create advanced lookup?",
        "How to search from bottom?"
    ],

    dataset:"employees",

    category:"Lookup",

    formula:"=XMATCH(H2,A:A)",

    syntax:"XMATCH(lookup_value,lookup_array,[match_mode],[search_mode])",

    explanation:"Finds the position of a value with more flexible matching options than MATCH.",

    breakdown:[
        "H2 = Lookup Value",
        "A:A = Lookup Range",
        "Returns Position"
    ],

    output:"5",

    business:"Used in modern Excel dashboards and dynamic models.",

    tips:[
        "Supports reverse search.",
        "Works better with dynamic arrays."
    ],

    errors:[
        "#N/A",
        "Value not found"
    ],

    alternative:"=MATCH(H2,A:A,0)"
},


// ==========================================================
// OFFSET
// ==========================================================
{
    title:"OFFSET",

    keywords:[
        "offset",
        "dynamic range",
        "dynamic reference",
        "move cell reference",
        "variable range",
        "advanced excel",
        "dashboard range",
        "dynamic chart",
        "rolling report",
        "data range expansion",
        "excel reference function",
        "automation",
        "mis dashboard",
        "financial model",
        "dynamic formula",
        "range calculation",
        "report automation",
        "table expansion",
        "business analysis"
    ],

    aliases:[
        "dynamic reference formula",
        "moving range formula",
        "variable cell reference",
        "dynamic data range"
    ],

    questions:[
        "How to create dynamic range?",
        "How to use OFFSET function?",
        "How to make dynamic charts?",
        "How to move cell reference?",
        "How to create rolling reports?"
    ],

    dataset:"sales",

    category:"Lookup",

    formula:"=OFFSET(A1,2,1)",

    syntax:"OFFSET(reference,rows,cols,[height],[width])",

    explanation:"Returns a reference shifted from a starting cell by specified rows and columns.",

    breakdown:[
        "A1 = Starting Cell",
        "2 = Move Rows",
        "1 = Move Columns"
    ],

    output:"Value from C3",

    business:"Used for dynamic dashboards, charts, and flexible reporting ranges.",

    tips:[
        "Useful for dynamic named ranges.",
        "Can create rolling reports."
    ],

    errors:[
        "#REF!",
        "Invalid reference"
    ],

    alternative:"=INDEX(A:C,3,3)"
},

// ==========================================================
// HLOOKUP
// ==========================================================
{
    title:"HLOOKUP",

    keywords:[
        "hlookup",
        "horizontal lookup",
        "lookup horizontal table",
        "find value in row",
        "return matching value",
        "excel lookup formula",
        "table lookup",
        "employee data lookup",
        "sales lookup",
        "price lookup",
        "old excel lookup",
        "row based lookup",
        "excel reference function",
        "data retrieval",
        "report automation",
        "advanced excel",
        "dashboard lookup",
        "business analysis",
        "master data search",
        "horizontal table search"
    ],

    aliases:[
        "horizontal vlookup",
        "row lookup formula",
        "lookup across rows",
        "table row search"
    ],

    questions:[
        "How to use HLOOKUP?",
        "How to lookup horizontally?",
        "How to find value from row?",
        "How to search horizontal table?",
        "What is horizontal lookup formula?"
    ],

    dataset:"sales",

    category:"Lookup",

    formula:"=HLOOKUP(H2,A1:F5,4,FALSE)",

    syntax:"HLOOKUP(value,table,row,FALSE)",

    explanation:"Searches the first row of a table and returns a value from a specified row.",

    breakdown:[
        "H2 = Lookup Value",
        "A1:F5 = Horizontal Table",
        "4 = Return Row Number"
    ],

    output:"50000",

    business:"Used for horizontal reports, monthly analysis, and comparison tables.",

    tips:[
        "Use XLOOKUP for modern Excel.",
        "Works best with horizontal data."
    ],

    errors:[
        "#N/A",
        "Wrong row number"
    ],

    alternative:"=XLOOKUP(H2,A1:F1,A4:F4)"
},


// ==========================================================
// VLOOKUP (Enhanced)
// ==========================================================
{
    title:"VLOOKUP",

    keywords:[
        "vlookup",
        "vertical lookup",
        "find employee salary",
        "lookup employee",
        "lookup value",
        "table lookup",
        "search value",
        "return matching value",
        "excel lookup formula",
        "old excel lookup",
        "column index number",
        "exact match lookup",
        "false match",
        "employee database",
        "salary report",
        "master data lookup",
        "data retrieval",
        "report automation",
        "advanced excel",
        "business analysis",
        "replace manual search",
        "lookup table"
    ],

    aliases:[
        "vertical lookup formula",
        "classic lookup",
        "excel search formula",
        "column lookup",
        "traditional lookup"
    ],

    questions:[
        "How to use VLOOKUP?",
        "How to find employee salary?",
        "How to lookup from table?",
        "How to search value vertically?",
        "How to replace manual lookup?"
    ],

    priorityRules:[
        "If query contains 'vlookup' => show VLOOKUP",
        "If query contains 'column index' => show VLOOKUP",
        "If query contains 'old excel' => show VLOOKUP",
        "If query contains 'dynamic lookup' => prefer XLOOKUP"
    ],

    dataset:"employees",

    category:"Lookup",

    formula:"=VLOOKUP(H2,A2:E100,4,FALSE)",

    syntax:"VLOOKUP(value,table,col,FALSE)",

    explanation:"Searches the first column of a table and returns matching data from another column.",

    breakdown:[
        "H2 = Employee ID",
        "A:E = Employee Table",
        "4 = Salary Column",
        "FALSE = Exact Match"
    ],

    output:"65000",

    business:"Used for employee salary, customer lookup, and master data reporting.",

    tips:[
        "Lookup value must be in first column.",
        "Use XLOOKUP for flexible searches."
    ],

    errors:[
        "#N/A",
        "#REF!",
        "Wrong column index"
    ],

    alternative:"=XLOOKUP(H2,A:A,D:D)"
},


// ==========================================================
// LOOKUP
// ==========================================================
{
    title:"LOOKUP",

    keywords:[
        "lookup",
        "find value",
        "search value",
        "return matching result",
        "old lookup function",
        "approximate lookup",
        "range lookup",
        "excel lookup formula",
        "data retrieval",
        "salary lookup",
        "price lookup",
        "commission calculation",
        "reference function",
        "advanced excel",
        "business analysis",
        "automation",
        "table search",
        "matching data",
        "report formula"
    ],

    aliases:[
        "old lookup formula",
        "range search formula",
        "approximate match lookup"
    ],

    questions:[
        "How to use LOOKUP?",
        "How to find approximate value?",
        "How to create range lookup?",
        "How to return matching data?"
    ],

    dataset:"commission",

    category:"Lookup",

    formula:"=LOOKUP(H2,A2:A10,B2:B10)",

    syntax:"LOOKUP(value,lookup_vector,result_vector)",

    explanation:"Searches a value in one range and returns corresponding value from another range.",

    breakdown:[
        "H2 = Search Value",
        "A2:A10 = Lookup Range",
        "B2:B10 = Result Range"
    ],

    output:"10%",

    business:"Used for slab calculations, commissions, and pricing rules.",

    tips:[
        "Data should be sorted ascending.",
        "XLOOKUP is preferred in modern Excel."
    ],

    errors:[
        "#N/A",
        "Unsorted data"
    ],

    alternative:"=XLOOKUP(H2,A:A,B:B,, -1)"
},


// ==========================================================
// CHOOSE
// ==========================================================
{
    title:"CHOOSE",

    keywords:[
        "choose",
        "select value",
        "index based selection",
        "option selection",
        "mapping formula",
        "category mapping",
        "status conversion",
        "employee grade",
        "month selection",
        "logical mapping",
        "excel reference function",
        "automation",
        "advanced excel",
        "dashboard formula",
        "business rules",
        "decision formula",
        "data classification",
        "report logic"
    ],

    aliases:[
        "value selector",
        "option formula",
        "index selection formula",
        "mapping function"
    ],

    questions:[
        "How to select value by number?",
        "How to map categories?",
        "How to use CHOOSE function?",
        "How to create option based result?"
    ],

    dataset:"employees",

    category:"Lookup",

    formula:'=CHOOSE(2,"Low","Medium","High")',

    syntax:"CHOOSE(index_num,value1,value2,...)",

    explanation:"Returns a value from a list based on index position.",

    breakdown:[
        "2 = Position Number",
        "List = Available Options",
        "Returns selected item"
    ],

    output:"Medium",

    business:"Used for mapping, selections, and dashboard controls.",

    tips:[
        "Useful for small fixed lists.",
        "Combine with MATCH for dynamic selection."
    ],

    errors:[
        "#VALUE!",
        "Invalid index"
    ],

    alternative:'=INDEX({"Low","Medium","High"},2)'
},


// ==========================================================
// INDIRECT
// ==========================================================
{
    title:"INDIRECT",

    keywords:[
        "indirect",
        "dynamic reference",
        "dynamic cell reference",
        "text reference",
        "change sheet reference",
        "dynamic formula",
        "multi sheet lookup",
        "variable reference",
        "excel reference function",
        "automation",
        "advanced excel",
        "dashboard formula",
        "dynamic reports",
        "financial model",
        "data retrieval",
        "sheet automation",
        "formula creation",
        "business reporting",
        "reference management"
    ],

    aliases:[
        "dynamic address formula",
        "text to reference",
        "flexible reference",
        "sheet reference formula"
    ],

    questions:[
        "How to create dynamic reference?",
        "How to reference another sheet?",
        "How to use INDIRECT?",
        "How to change formula reference dynamically?",
        "How to create dynamic reports?"
    ],

    dataset:"monthly_reports",

    category:"Lookup",

    formula:'=INDIRECT("Sheet2!A"&B2)',

    syntax:"INDIRECT(ref_text)",

    explanation:"Converts a text string into a valid cell reference.",

    breakdown:[
        "Sheet2!A = Target Reference",
        "B2 = Dynamic Row Number",
        "Creates Cell Reference"
    ],

    output:"Sales Value",

    business:"Used for dynamic dashboards and multi-sheet reporting.",

    tips:[
        "Avoid excessive use in large files.",
        "Can slow workbook performance."
    ],

    errors:[
        "#REF!",
        "Invalid reference text"
    ],

    alternative:"=INDEX(Sheet2!A:A,B2)"
},

// ==========================================================
// TODAY
// ==========================================================
{
    title:"TODAY",

    keywords:[
        "today",
        "current date",
        "today date formula",
        "automatic date",
        "live date",
        "system date",
        "date tracking",
        "deadline calculation",
        "due date",
        "employee age calculation",
        "tenure calculation",
        "sla tracking",
        "mis report",
        "daily report",
        "excel date function",
        "date automation",
        "dashboard date",
        "aging report",
        "compliance tracking",
        "advanced excel"
    ],

    aliases:[
        "current day formula",
        "auto date formula",
        "dynamic date",
        "live date function"
    ],

    questions:[
        "How to get today's date?",
        "How to insert current date automatically?",
        "How to calculate days from today?",
        "How to use TODAY function?",
        "How to track due dates?"
    ],

    dataset:"tasks",

    category:"Date",

    formula:"=TODAY()",

    syntax:"TODAY()",

    explanation:"Returns the current date automatically based on system date.",

    breakdown:[
        "No arguments required",
        "Returns today's date",
        "Updates automatically"
    ],

    output:"07/28/2026",

    business:"Used for aging reports, SLA tracking, deadlines, and daily MIS.",

    tips:[
        "Updates whenever workbook recalculates.",
        "Use with DATEDIF for age or tenure."
    ],

    errors:[
        "No common errors"
    ],

    alternative:"=NOW()"
},


// ==========================================================
// NOW
// ==========================================================
{
    title:"NOW",

    keywords:[
        "now",
        "current date time",
        "live timestamp",
        "current time",
        "date and time",
        "timestamp formula",
        "transaction time",
        "log time",
        "audit tracking",
        "activity monitoring",
        "excel date function",
        "automation",
        "report timestamp",
        "system time",
        "dashboard refresh",
        "process tracking",
        "sla monitoring",
        "advanced excel",
        "compliance report",
        "time tracking"
    ],

    aliases:[
        "timestamp formula",
        "live date time",
        "current datetime",
        "automatic timestamp"
    ],

    questions:[
        "How to get current date and time?",
        "How to create timestamp?",
        "How to track transaction time?",
        "How to use NOW function?",
        "How to record refresh time?"
    ],

    dataset:"transactions",

    category:"Date",

    formula:"=NOW()",

    syntax:"NOW()",

    explanation:"Returns the current date and time.",

    breakdown:[
        "No arguments required",
        "Returns current date",
        "Returns current time"
    ],

    output:"07/28/2026 08:30 PM",

    business:"Used for audit logs, transaction tracking, and reporting timestamps.",

    tips:[
        "Updates automatically.",
        "Useful in dashboards."
    ],

    errors:[
        "No common errors"
    ],

    alternative:"=TODAY()"
},


// ==========================================================
// DATE
// ==========================================================
{
    title:"DATE",

    keywords:[
        "date",
        "create date",
        "combine year month day",
        "generate date",
        "date creation",
        "dynamic date",
        "date calculation",
        "joining date",
        "expiry date",
        "due date",
        "excel date function",
        "automation",
        "financial reporting",
        "mis report",
        "date conversion",
        "calendar formula",
        "schedule tracking",
        "advanced excel",
        "business date"
    ],

    aliases:[
        "date generator",
        "build date formula",
        "create calendar date",
        "combine date values"
    ],

    questions:[
        "How to create date from year month day?",
        "How to combine date values?",
        "How to use DATE function?",
        "How to generate due date?",
        "How to create dynamic dates?"
    ],

    dataset:"employees",

    category:"Date",

    formula:"=DATE(A2,B2,C2)",

    syntax:"DATE(year,month,day)",

    explanation:"Creates a valid Excel date from separate year, month, and day values.",

    breakdown:[
        "A2 = Year",
        "B2 = Month",
        "C2 = Day",
        "Creates Complete Date"
    ],

    output:"07/28/2026",

    business:"Used for employee joining dates, expiry tracking, and reporting calendars.",

    tips:[
        "Useful when dates come from separate columns.",
        "Works well with YEAR/MONTH/DAY."
    ],

    errors:[
        "#VALUE!",
        "Invalid date values"
    ],

    alternative:"=TODAY()"
},


// ==========================================================
// YEAR
// ==========================================================
{
    title:"YEAR",

    keywords:[
        "year",
        "extract year",
        "get year from date",
        "joining year",
        "financial year",
        "date analysis",
        "year wise report",
        "annual report",
        "employee tenure",
        "trend analysis",
        "excel date function",
        "mis reporting",
        "dashboard analysis",
        "date extraction",
        "automation",
        "advanced excel",
        "business analytics",
        "calendar analysis",
        "period reporting",
        "date breakdown"
    ],

    aliases:[
        "year extraction",
        "date year formula",
        "annual grouping formula"
    ],

    questions:[
        "How to extract year from date?",
        "How to find joining year?",
        "How to create year wise report?",
        "How to use YEAR function?",
        "How to group dates by year?"
    ],

    dataset:"employees",

    category:"Date",

    formula:"=YEAR(B2)",

    syntax:"YEAR(serial_number)",

    explanation:"Returns the year component from a date.",

    breakdown:[
        "B2 = Date Value",
        "Extracts Year",
        "Returns Four Digit Year"
    ],

    output:"2026",

    business:"Used for yearly MIS, employee analysis, and financial reporting.",

    tips:[
        "Combine with MONTH for period analysis.",
        "Useful in Pivot reports."
    ],

    errors:[
        "#VALUE!",
        "Invalid date"
    ],

    alternative:"=TEXT(B2,\"YYYY\")"
},


// ==========================================================
// MONTH
// ==========================================================
{
    title:"MONTH",

    keywords:[
        "month",
        "extract month",
        "get month from date",
        "monthly report",
        "month wise analysis",
        "date breakdown",
        "financial month",
        "employee joining month",
        "sales month",
        "MIS monthly report",
        "excel date function",
        "dashboard analysis",
        "automation",
        "advanced excel",
        "period analysis",
        "calendar reporting",
        "trend analysis",
        "date extraction",
        "business analytics",
        "monthly grouping"
    ],

    aliases:[
        "month extraction",
        "date month formula",
        "monthly grouping formula"
    ],

    questions:[
        "How to extract month from date?",
        "How to create monthly report?",
        "How to find joining month?",
        "How to use MONTH function?",
        "How to group data by month?"
    ],

    dataset:"sales",

    category:"Date",

    formula:"=MONTH(B2)",

    syntax:"MONTH(serial_number)",

    explanation:"Returns the month number from a date.",

    breakdown:[
        "B2 = Date Value",
        "Extracts Month",
        "Returns Month Number"
    ],

    output:"7",

    business:"Used for monthly MIS, sales analysis, and trend reporting.",

    tips:[
        "Combine with TEXT for month names.",
        "Useful in PivotTables."
    ],

    errors:[
        "#VALUE!",
        "Invalid date"
    ],

    alternative:'=TEXT(B2,"MMMM")'
},

// ==========================================================
// DAY
// ==========================================================
{
    title:"DAY",

    keywords:[
        "day",
        "extract day",
        "get day from date",
        "date day formula",
        "day number",
        "transaction day",
        "joining date day",
        "date breakdown",
        "calendar analysis",
        "daily report",
        "excel date function",
        "date extraction",
        "mis reporting",
        "dashboard analysis",
        "automation",
        "advanced excel",
        "business reporting",
        "period analysis",
        "date grouping",
        "schedule tracking"
    ],

    aliases:[
        "day extraction",
        "date day finder",
        "extract date number"
    ],

    questions:[
        "How to extract day from date?",
        "How to find transaction day?",
        "How to use DAY function?",
        "How to get date number?",
        "How to separate day from date?"
    ],

    dataset:"transactions",

    category:"Date",

    formula:"=DAY(B2)",

    syntax:"DAY(serial_number)",

    explanation:"Returns the day number from a date.",

    breakdown:[
        "B2 = Date Value",
        "Extracts Day",
        "Returns Day Number"
    ],

    output:"28",

    business:"Used for daily reports, transaction analysis, and date breakdown.",

    tips:[
        "Combine YEAR, MONTH, DAY for complete date analysis.",
        "Useful in calendar reports."
    ],

    errors:[
        "#VALUE!",
        "Invalid date"
    ],

    alternative:'=TEXT(B2,"DD")'
},


// ==========================================================
// DAYS
// ==========================================================
{
    title:"DAYS",

    keywords:[
        "days",
        "difference between dates",
        "calculate days",
        "date difference",
        "aging days",
        "overdue days",
        "sla days",
        "turnaround time",
        "payment delay",
        "invoice aging",
        "excel date function",
        "date calculation",
        "automation",
        "mis reporting",
        "compliance tracking",
        "business days",
        "deadline calculation",
        "project duration",
        "advanced excel",
        "time analysis"
    ],

    aliases:[
        "date difference formula",
        "number of days formula",
        "days calculator",
        "date gap formula"
    ],

    questions:[
        "How to calculate days between dates?",
        "How to find overdue days?",
        "How to calculate SLA days?",
        "How to use DAYS function?",
        "How many days between two dates?"
    ],

    dataset:"tasks",

    category:"Date",

    formula:"=DAYS(C2,B2)",

    syntax:"DAYS(end_date,start_date)",

    explanation:"Returns the number of days between two dates.",

    breakdown:[
        "C2 = End Date",
        "B2 = Start Date",
        "Returns Difference"
    ],

    output:"45",

    business:"Used for aging reports, SLA monitoring, and turnaround analysis.",

    tips:[
        "Use NETWORKDAYS for working days.",
        "Useful for compliance tracking."
    ],

    errors:[
        "#VALUE!",
        "Invalid date"
    ],

    alternative:"=C2-B2"
},


// ==========================================================
// DATEDIF
// ==========================================================
{
    title:"DATEDIF",

    keywords:[
        "datedif",
        "date difference",
        "calculate age",
        "employee age",
        "experience calculation",
        "tenure calculation",
        "years months days",
        "service period",
        "joining experience",
        "retirement calculation",
        "excel date function",
        "hr analytics",
        "employee report",
        "duration formula",
        "automation",
        "advanced excel",
        "time calculation",
        "business reporting",
        "date analysis",
        "period calculation"
    ],

    aliases:[
        "age calculator",
        "experience formula",
        "tenure formula",
        "years between dates"
    ],

    questions:[
        "How to calculate age?",
        "How to calculate employee experience?",
        "How to find years between dates?",
        "How to use DATEDIF?",
        "How to calculate service period?"
    ],

    dataset:"employees",

    category:"Date",

    formula:'=DATEDIF(B2,TODAY(),"Y")',

    syntax:"DATEDIF(start_date,end_date,unit)",

    explanation:"Calculates the difference between two dates in years, months, or days.",

    breakdown:[
        "B2 = Joining Date",
        "TODAY() = Current Date",
        "\"Y\" = Complete Years"
    ],

    output:"8",

    business:"Used for employee tenure, age calculation, and HR analytics.",

    tips:[
        "Hidden Excel function but widely used.",
        "Units: Y, M, D."
    ],

    errors:[
        "#NUM!",
        "Start date greater than end date"
    ],

    alternative:'=YEARFRAC(B2,TODAY())'
},


// ==========================================================
// EDATE
// ==========================================================
{
    title:"EDATE",

    keywords:[
        "edate",
        "add months",
        "subtract months",
        "future date",
        "previous date",
        "maturity date",
        "renewal date",
        "expiry date",
        "contract period",
        "loan calculation",
        "subscription date",
        "excel date function",
        "date automation",
        "financial calculation",
        "schedule planning",
        "advanced excel",
        "business reporting",
        "calendar calculation",
        "timeline formula",
        "period management"
    ],

    aliases:[
        "month addition formula",
        "future month date",
        "date offset formula"
    ],

    questions:[
        "How to add months to date?",
        "How to calculate maturity date?",
        "How to extend contract date?",
        "How to use EDATE function?",
        "How to find future date?"
    ],

    dataset:"contracts",

    category:"Date",

    formula:"=EDATE(B2,12)",

    syntax:"EDATE(start_date,months)",

    explanation:"Returns a date after adding or subtracting specified months.",

    breakdown:[
        "B2 = Start Date",
        "12 = Months Added",
        "Returns Future Date"
    ],

    output:"07/28/2027",

    business:"Used for contracts, loans, renewals, and subscription tracking.",

    tips:[
        "Negative months calculate previous dates.",
        "Useful in financial models."
    ],

    errors:[
        "#VALUE!",
        "Invalid date"
    ],

    alternative:"=DATE(YEAR(B2)+1,MONTH(B2),DAY(B2))"
},


// ==========================================================
// EOMONTH
// ==========================================================
{
    title:"EOMONTH",

    keywords:[
        "eomonth",
        "month end date",
        "last day of month",
        "month closing",
        "financial month end",
        "reporting date",
        "period end",
        "account closing",
        "monthly MIS",
        "finance report",
        "excel date function",
        "calendar formula",
        "automation",
        "advanced excel",
        "dashboard date",
        "business reporting",
        "month closing process",
        "period calculation",
        "financial analysis",
        "date management"
    ],

    aliases:[
        "month end formula",
        "closing date formula",
        "last date of month",
        "period end formula"
    ],

    questions:[
        "How to find month end date?",
        "How to calculate closing date?",
        "How to get last day of month?",
        "How to use EOMONTH function?",
        "How to create monthly reports?"
    ],

    dataset:"finance",

    category:"Date",

    formula:"=EOMONTH(B2,0)",

    syntax:"EOMONTH(start_date,months)",

    explanation:"Returns the last day of a month before or after a specified number of months.",

    breakdown:[
        "B2 = Starting Date",
        "0 = Same Month",
        "Returns Month End Date"
    ],

    output:"07/31/2026",

    business:"Used for month-end closing, MIS reporting, and financial analysis.",

    tips:[
        "Use EOMONTH(B2,-1) for previous month end.",
        "Useful for financial calendars."
    ],

    errors:[
        "#VALUE!",
        "Invalid date"
    ],

    alternative:"=DATE(YEAR(B2),MONTH(B2)+1,0)"
},

// ==========================================================
// NETWORKDAYS
// ==========================================================
{
    title:"NETWORKDAYS",

    keywords:[
        "networkdays",
        "working days",
        "business days",
        "calculate working days",
        "exclude weekends",
        "exclude holidays",
        "sla calculation",
        "tat calculation",
        "turnaround time",
        "employee attendance",
        "leave calculation",
        "project duration",
        "excel date function",
        "banking working days",
        "compliance timeline",
        "mis report",
        "hr analytics",
        "deadline calculation",
        "business calendar",
        "advanced excel"
    ],

    aliases:[
        "workday counter",
        "business day calculator",
        "office days formula",
        "weekday difference"
    ],

    questions:[
        "How to calculate working days?",
        "How to exclude weekends?",
        "How to calculate SLA days?",
        "How to count office days?",
        "How to use NETWORKDAYS function?"
    ],

    dataset:"tasks",

    category:"Date",

    formula:"=NETWORKDAYS(B2,C2)",

    syntax:"NETWORKDAYS(start_date,end_date,[holidays])",

    explanation:"Returns the number of working days between two dates excluding weekends.",

    breakdown:[
        "B2 = Start Date",
        "C2 = End Date",
        "Excludes Saturday and Sunday"
    ],

    output:"22",

    business:"Used for SLA monitoring, leave calculation, and working day reports.",

    tips:[
        "Add holiday range to exclude public holidays.",
        "Commonly used in HR and compliance reports."
    ],

    errors:[
        "#VALUE!",
        "Invalid date"
    ],

    alternative:"=DAYS(C2,B2)-INT(DAYS(C2,B2)/7)*2"
},


// ==========================================================
// NETWORKDAYS.INTL
// ==========================================================
{
    title:"NETWORKDAYS.INTL",

    keywords:[
        "networkdays intl",
        "custom working days",
        "different weekends",
        "business days custom",
        "exclude friday saturday",
        "custom calendar",
        "international working days",
        "regional holidays",
        "global sla",
        "working day calculation",
        "employee shift",
        "attendance calculation",
        "excel date function",
        "automation",
        "advanced excel",
        "business reporting",
        "project planning",
        "compliance tracking",
        "custom weekend",
        "date analysis"
    ],

    aliases:[
        "custom workday formula",
        "regional business days",
        "flexible working days"
    ],

    questions:[
        "How to calculate custom working days?",
        "How to exclude Friday Saturday?",
        "How to create regional calendars?",
        "How to use NETWORKDAYS.INTL?",
        "How to calculate global SLA?"
    ],

    dataset:"global_projects",

    category:"Date",

    formula:'=NETWORKDAYS.INTL(B2,C2,"0000011")',

    syntax:"NETWORKDAYS.INTL(start_date,end_date,[weekend],[holidays])",

    explanation:"Calculates working days using custom weekend patterns.",

    breakdown:[
        "B2 = Start Date",
        "C2 = End Date",
        "0000011 = Saturday Sunday Weekend"
    ],

    output:"20",

    business:"Used for global teams with different working calendars.",

    tips:[
        "Useful for international operations.",
        "Weekend pattern can be customized."
    ],

    errors:[
        "#VALUE!"
    ],

    alternative:"=NETWORKDAYS(B2,C2)"
},


// ==========================================================
// WORKDAY
// ==========================================================
{
    title:"WORKDAY",

    keywords:[
        "workday",
        "future working date",
        "add business days",
        "calculate deadline",
        "due date",
        "sla deadline",
        "project completion date",
        "exclude weekends",
        "office day calculation",
        "delivery date",
        "excel date function",
        "task planning",
        "compliance deadline",
        "automation",
        "advanced excel",
        "business calendar",
        "reporting timeline",
        "process tracking",
        "turnaround time"
    ],

    aliases:[
        "business date formula",
        "deadline calculator",
        "working date finder"
    ],

    questions:[
        "How to add working days?",
        "How to calculate due date?",
        "How to find project completion date?",
        "How to use WORKDAY function?",
        "How to calculate SLA deadline?"
    ],

    dataset:"projects",

    category:"Date",

    formula:"=WORKDAY(B2,10)",

    syntax:"WORKDAY(start_date,days,[holidays])",

    explanation:"Returns a future or past date after adding working days.",

    breakdown:[
        "B2 = Start Date",
        "10 = Working Days Added",
        "Returns Deadline Date"
    ],

    output:"08/11/2026",

    business:"Used for SLA deadlines, project planning, and delivery tracking.",

    tips:[
        "Excludes weekends automatically.",
        "Add holidays for accurate results."
    ],

    errors:[
        "#VALUE!",
        "Invalid date"
    ],

    alternative:"=B2+10"
},


// ==========================================================
// WORKDAY.INTL
// ==========================================================
{
    title:"WORKDAY.INTL",

    keywords:[
        "workday intl",
        "custom workday",
        "international calendar",
        "custom weekend",
        "shift working days",
        "global deadline",
        "regional calendar",
        "different weekends",
        "business date",
        "sla calculation",
        "project timeline",
        "excel date function",
        "automation",
        "advanced excel",
        "working schedule",
        "compliance deadline",
        "global reporting",
        "operations planning",
        "date management"
    ],

    aliases:[
        "custom deadline formula",
        "flexible workday",
        "regional work calendar"
    ],

    questions:[
        "How to calculate custom business date?",
        "How to exclude different weekends?",
        "How to use WORKDAY.INTL?",
        "How to create global calendar?",
        "How to calculate international SLA?"
    ],

    dataset:"global_operations",

    category:"Date",

    formula:'=WORKDAY.INTL(B2,15,"0000011")',

    syntax:"WORKDAY.INTL(start_date,days,[weekend],[holidays])",

    explanation:"Returns a date after adding working days based on custom weekend rules.",

    breakdown:[
        "B2 = Start Date",
        "15 = Working Days",
        "Weekend Pattern Applied"
    ],

    output:"08/18/2026",

    business:"Used for international teams and customized working calendars.",

    tips:[
        "Useful for different country calendars.",
        "Supports custom weekends."
    ],

    errors:[
        "#VALUE!"
    ],

    alternative:"=WORKDAY(B2,15)"
},


// ==========================================================
// WEEKDAY
// ==========================================================
{
    title:"WEEKDAY",

    keywords:[
        "weekday",
        "day name",
        "day number",
        "find weekday",
        "monday tuesday formula",
        "calendar analysis",
        "date classification",
        "weekly report",
        "attendance analysis",
        "shift planning",
        "excel date function",
        "automation",
        "advanced excel",
        "schedule management",
        "date reporting",
        "business calendar",
        "employee attendance",
        "weekly trend",
        "time analysis",
        "date extraction"
    ],

    aliases:[
        "day finder",
        "week day formula",
        "date day name",
        "calendar day"
    ],

    questions:[
        "How to find day name from date?",
        "How to identify weekday?",
        "How to check weekend?",
        "How to use WEEKDAY function?",
        "How to create weekly reports?"
    ],

    dataset:"attendance",

    category:"Date",

    formula:"=WEEKDAY(B2,2)",

    syntax:"WEEKDAY(serial_number,[return_type])",

    explanation:"Returns the day number of the week from a date.",

    breakdown:[
        "B2 = Date",
        "2 = Monday as first day",
        "Returns weekday number"
    ],

    output:"2",

    business:"Used for attendance analysis, shift planning, and weekly reports.",

    tips:[
        "Use TEXT(B2,\"dddd\") to get day name.",
        "Useful for weekend detection."
    ],

    errors:[
        "#VALUE!",
        "Invalid date"
    ],

    alternative:'=TEXT(B2,"dddd")'
},

// ==========================================================
// IF
// ==========================================================
{
    title:"IF",

    keywords:[
        "if",
        "if condition",
        "logical test",
        "condition formula",
        "check condition",
        "true false formula",
        "pass fail formula",
        "status formula",
        "decision formula",
        "employee status",
        "salary condition",
        "approval formula",
        "excel logical function",
        "automation",
        "advanced excel",
        "business rules",
        "data validation",
        "report logic",
        "classification formula",
        "mis reporting"
    ],

    aliases:[
        "condition formula",
        "decision making formula",
        "yes no formula",
        "comparison formula"
    ],

    questions:[
        "How to use IF formula?",
        "How to create condition?",
        "How to check pass fail?",
        "How to return yes or no?",
        "How to apply logical condition?"
    ],

    priorityRules:[
        "If query contains only 'if' => show IF",
        "If query contains 'multiple conditions' => prefer IFS",
        "If query contains 'nested condition' => show IF"
    ],

    dataset:"employees",

    category:"Logical",

    formula:'=IF(C2>=50000,"High","Low")',

    syntax:"IF(logical_test,value_if_true,value_if_false)",

    explanation:"Checks a condition and returns one value if TRUE and another value if FALSE.",

    breakdown:[
        "C2 = Salary Value",
        ">=50000 = Condition",
        "High/Low = Result"
    ],

    output:"High",

    business:"Used for employee grading, approvals, status tracking, and reporting logic.",

    tips:[
        "Can be nested for multiple conditions.",
        "Use IFS for cleaner multiple conditions."
    ],

    errors:[
        "#VALUE!",
        "Incorrect logical test"
    ],

    alternative:'=IFS(C2>=80000,"High",C2>=50000,"Medium",TRUE,"Low")'
},


// ==========================================================
// IFS
// ==========================================================
{
    title:"IFS",

    keywords:[
        "ifs",
        "multiple conditions",
        "multiple if",
        "many conditions",
        "nested if alternative",
        "multiple criteria formula",
        "grade calculation",
        "salary band",
        "category formula",
        "logical conditions",
        "excel logical function",
        "decision tree",
        "automation",
        "advanced excel",
        "business rules",
        "classification",
        "score calculation",
        "performance rating",
        "mis reporting",
        "data analysis"
    ],

    aliases:[
        "multiple if formula",
        "condition list formula",
        "advanced if formula",
        "nested if replacement"
    ],

    questions:[
        "How to use multiple conditions?",
        "How to replace nested IF?",
        "How to create multiple criteria?",
        "How to calculate grades?",
        "How to use IFS function?"
    ],

    priorityRules:[
        "If query contains 'multiple conditions' => show IFS",
        "If query contains 'many if' => show IFS",
        "If query contains simple if => show IF"
    ],

    dataset:"performance",

    category:"Logical",

    formula:'=IFS(B2>=90,"Excellent",B2>=75,"Good",B2>=50,"Average",TRUE,"Poor")',

    syntax:"IFS(test1,result1,test2,result2,...)",

    explanation:"Checks multiple conditions and returns the result for the first TRUE condition.",

    breakdown:[
        "B2 = Score",
        "Multiple Conditions Applied",
        "Returns Matching Category"
    ],

    output:"Good",

    business:"Used for grading, salary bands, ratings, and classification.",

    tips:[
        "Cleaner than multiple nested IF formulas.",
        "Available in modern Excel."
    ],

    errors:[
        "#N/A",
        "No TRUE condition"
    ],

    alternative:'=IF(B2>=90,"Excellent",IF(B2>=75,"Good","Average"))'
},


// ==========================================================
// AND
// ==========================================================
{
    title:"AND",

    keywords:[
        "and",
        "all conditions",
        "multiple criteria",
        "check all conditions",
        "combined condition",
        "logical test",
        "approval criteria",
        "eligibility check",
        "employee validation",
        "excel logical function",
        "automation",
        "advanced excel",
        "business rules",
        "data validation",
        "decision formula",
        "compliance check",
        "control testing",
        "risk assessment",
        "report logic",
        "condition checking"
    ],

    aliases:[
        "all criteria formula",
        "multiple test formula",
        "combined condition"
    ],

    questions:[
        "How to check all conditions?",
        "How to use AND formula?",
        "How to validate multiple criteria?",
        "How to combine conditions?",
        "How to check eligibility?"
    ],

    dataset:"employees",

    category:"Logical",

    formula:"=AND(B2>=18,C2=\"Active\")",

    syntax:"AND(logical1,logical2,...)",

    explanation:"Returns TRUE only when all conditions are TRUE.",

    breakdown:[
        "B2>=18 = Age Check",
        "C2 Active = Status Check",
        "Both must be TRUE"
    ],

    output:"TRUE",

    business:"Used for approvals, eligibility checks, and compliance controls.",

    tips:[
        "Combine with IF for final result.",
        "Useful in validation rules."
    ],

    errors:[
        "#VALUE!"
    ],

    alternative:'=IF(AND(B2>=18,C2="Active"),"Yes","No")'
},


// ==========================================================
// OR
// ==========================================================
{
    title:"OR",

    keywords:[
        "or",
        "any condition",
        "one condition",
        "multiple options",
        "alternative criteria",
        "logical test",
        "approval condition",
        "exception check",
        "risk identification",
        "excel logical function",
        "automation",
        "advanced excel",
        "business rules",
        "data validation",
        "decision formula",
        "compliance testing",
        "error checking",
        "report logic",
        "condition analysis"
    ],

    aliases:[
        "either condition formula",
        "option condition",
        "alternative criteria"
    ],

    questions:[
        "How to check any condition?",
        "How to use OR formula?",
        "How to check multiple options?",
        "How to apply alternative criteria?",
        "How to find exceptions?"
    ],

    dataset:"transactions",

    category:"Logical",

    formula:'=OR(B2="Pending",B2="Rejected")',

    syntax:"OR(logical1,logical2,...)",

    explanation:"Returns TRUE if at least one condition is TRUE.",

    breakdown:[
        "Condition 1 = Pending",
        "Condition 2 = Rejected",
        "Any TRUE returns TRUE"
    ],

    output:"TRUE",

    business:"Used for exception reporting, approval checks, and alerts.",

    tips:[
        "Combine with IF for meaningful output.",
        "Useful for multiple scenarios."
    ],

    errors:[
        "#VALUE!"
    ],

    alternative:'=IF(OR(B2="Pending",B2="Rejected"),"Review","OK")'
},


// ==========================================================
// NOT
// ==========================================================
{
    title:"NOT",

    keywords:[
        "not",
        "reverse condition",
        "opposite logical result",
        "negative condition",
        "exclude condition",
        "check not equal",
        "exception formula",
        "logical reversal",
        "validation formula",
        "excel logical function",
        "automation",
        "advanced excel",
        "business rules",
        "data filtering",
        "error checking",
        "compliance check",
        "report logic",
        "condition analysis",
        "decision formula"
    ],

    aliases:[
        "reverse logic formula",
        "opposite condition",
        "negative test"
    ],

    questions:[
        "How to reverse condition?",
        "How to use NOT function?",
        "How to check not equal?",
        "How to exclude values?",
        "How to create negative logic?"
    ],

    dataset:"employees",

    category:"Logical",

    formula:'=NOT(B2="Inactive")',

    syntax:"NOT(logical)",

    explanation:"Reverses the logical result. TRUE becomes FALSE and FALSE becomes TRUE.",

    breakdown:[
        "B2 = Status",
        "Checks Inactive",
        "Returns Opposite Result"
    ],

    output:"TRUE",

    business:"Used for exclusion rules, validation, and exception checks.",

    tips:[
        "Useful with IF and AND.",
        "Simplifies negative conditions."
    ],

    errors:[
        "#VALUE!"
    ],

    alternative:'=B2<>"Inactive"'
},

// ==========================================================
// IFERROR
// ==========================================================
{
    title:"IFERROR",

    keywords:[
        "iferror",
        "error handling",
        "remove error",
        "replace error message",
        "#value error",
        "#div/0 error",
        "#n/a error",
        "formula error",
        "avoid error",
        "custom error message",
        "error management",
        "excel logical function",
        "data cleaning",
        "report preparation",
        "automation",
        "advanced excel",
        "dashboard formula",
        "business reporting",
        "formula validation",
        "exception handling"
    ],

    aliases:[
        "error catcher",
        "error replacement formula",
        "safe formula",
        "error handling function"
    ],

    questions:[
        "How to remove formula errors?",
        "How to replace error with text?",
        "How to use IFERROR?",
        "How to handle division errors?",
        "How to show custom error message?"
    ],

    priorityRules:[
        "If query contains #N/A specifically => prefer IFNA",
        "If query contains general error => show IFERROR"
    ],

    dataset:"sales",

    category:"Logical",

    formula:'=IFERROR(A2/B2,"Not Available")',

    syntax:"IFERROR(value,value_if_error)",

    explanation:"Returns a custom result when a formula produces an error.",

    breakdown:[
        "A2/B2 = Formula Calculation",
        "\"Not Available\" = Error Result",
        "Returns Clean Output"
    ],

    output:"Not Available",

    business:"Used in dashboards, reports, and error-free presentations.",

    tips:[
        "Use after lookup formulas.",
        "Avoid hiding genuine data issues."
    ],

    errors:[
        "#VALUE!",
        "#DIV/0!"
    ],

    alternative:'=IF(ISERROR(A2/B2),"Not Available",A2/B2)'
},


// ==========================================================
// IFNA
// ==========================================================
{
    title:"IFNA",

    keywords:[
        "ifna",
        "#n/a error",
        "lookup error",
        "missing value",
        "vlookup error",
        "xlookup error",
        "not found message",
        "replace n/a",
        "lookup error handling",
        "excel logical function",
        "data validation",
        "master data",
        "automation",
        "advanced excel",
        "report preparation",
        "dashboard lookup",
        "business analysis",
        "exception handling",
        "formula cleaning"
    ],

    aliases:[
        "n/a error handler",
        "lookup error remover",
        "missing value handler"
    ],

    questions:[
        "How to remove #N/A error?",
        "How to handle lookup missing value?",
        "How to use IFNA?",
        "How to replace not found result?",
        "How to fix XLOOKUP error?"
    ],

    dataset:"employees",

    category:"Logical",

    formula:'=IFNA(XLOOKUP(H2,A:A,D:D),"Not Found")',

    syntax:"IFNA(value,value_if_na)",

    explanation:"Handles only #N/A errors and returns a custom value.",

    breakdown:[
        "XLOOKUP = Lookup Formula",
        "Not Found = If #N/A occurs",
        "Returns Clean Result"
    ],

    output:"Not Found",

    business:"Used for lookup reports and missing data handling.",

    tips:[
        "Prefer IFNA for lookup-specific errors.",
        "Does not hide other errors."
    ],

    errors:[
        "#N/A"
    ],

    alternative:'=IFERROR(XLOOKUP(H2,A:A,D:D),"Not Found")'
},


// ==========================================================
// SWITCH
// ==========================================================
{
    title:"SWITCH",

    keywords:[
        "switch",
        "multiple value mapping",
        "replace nested if",
        "condition mapping",
        "status mapping",
        "category conversion",
        "grade mapping",
        "employee rating",
        "logical mapping",
        "multiple output formula",
        "excel logical function",
        "automation",
        "advanced excel",
        "business rules",
        "classification formula",
        "dashboard formula",
        "data transformation",
        "report logic",
        "decision formula",
        "workflow automation"
    ],

    aliases:[
        "mapping formula",
        "nested if alternative",
        "value based condition",
        "category selector"
    ],

    questions:[
        "How to replace nested IF?",
        "How to map values?",
        "How to create status mapping?",
        "How to use SWITCH function?",
        "How to convert codes into names?"
    ],

    dataset:"employees",

    category:"Logical",

    formula:'=SWITCH(B2,"A","Excellent","B","Good","C","Average","Other")',

    syntax:"SWITCH(expression,value1,result1,...,default)",

    explanation:"Compares one value against multiple options and returns the matching result.",

    breakdown:[
        "B2 = Grade Code",
        "A/B/C = Matching Values",
        "Returns Category"
    ],

    output:"Excellent",

    business:"Used for status mapping, grades, and reporting categories.",

    tips:[
        "Cleaner than multiple IF conditions.",
        "Best for exact matching."
    ],

    errors:[
        "#N/A",
        "No matching value"
    ],

    alternative:'=IFS(B2="A","Excellent",B2="B","Good")'
},


// ==========================================================
// TRUE
// ==========================================================
{
    title:"TRUE",

    keywords:[
        "true",
        "true value",
        "logical true",
        "always true",
        "condition default",
        "logical test",
        "formula validation",
        "if formula default",
        "excel logical function",
        "automation",
        "advanced excel",
        "business rules",
        "decision formula",
        "comparison result",
        "validation check",
        "report logic"
    ],

    aliases:[
        "true function",
        "logical yes",
        "default true value"
    ],

    questions:[
        "How to return TRUE?",
        "How to use TRUE function?",
        "How to create default condition?",
        "How to check logical result?"
    ],

    dataset:"validation",

    category:"Logical",

    formula:"=TRUE()",

    syntax:"TRUE()",

    explanation:"Returns the logical value TRUE.",

    breakdown:[
        "No arguments required",
        "Returns TRUE",
        "Used in logical formulas"
    ],

    output:"TRUE",

    business:"Used as default conditions in logical formulas.",

    tips:[
        "Commonly used with IFS as final condition.",
        "Can also be written as TRUE."
    ],

    errors:[
        "No common errors"
    ],

    alternative:"=1=1"
},


// ==========================================================
// FALSE
// ==========================================================
{
    title:"FALSE",

    keywords:[
        "false",
        "false value",
        "logical false",
        "always false",
        "negative condition",
        "default false",
        "logical test",
        "formula validation",
        "excel logical function",
        "automation",
        "advanced excel",
        "business rules",
        "decision formula",
        "comparison result",
        "validation check",
        "report logic"
    ],

    aliases:[
        "false function",
        "logical no",
        "default false value"
    ],

    questions:[
        "How to return FALSE?",
        "How to use FALSE function?",
        "How to create false condition?",
        "How to test logical results?"
    ],

    dataset:"validation",

    category:"Logical",

    formula:"=FALSE()",

    syntax:"FALSE()",

    explanation:"Returns the logical value FALSE.",

    breakdown:[
        "No arguments required",
        "Returns FALSE",
        "Used in logical tests"
    ],

    output:"FALSE",

    business:"Used for testing, validation, and logical comparisons.",

    tips:[
        "Useful in formulas requiring a FALSE condition.",
        "Can also be written as FALSE."
    ],

    errors:[
        "No common errors"
    ],

    alternative:"=1=2"
},

// ==========================================================
// SUM
// ==========================================================
{
    title:"SUM",

    keywords:[
        "sum",
        "add values",
        "total",
        "calculate total",
        "addition formula",
        "grand total",
        "column total",
        "row total",
        "expense total",
        "sales total",
        "salary total",
        "excel math function",
        "calculation formula",
        "mis report",
        "financial report",
        "data analysis",
        "automation",
        "advanced excel",
        "dashboard total",
        "business calculation"
    ],

    aliases:[
        "total formula",
        "addition function",
        "aggregate total",
        "numeric addition"
    ],

    questions:[
        "How to calculate total?",
        "How to add numbers?",
        "How to sum a column?",
        "How to use SUM function?",
        "How to create grand total?"
    ],

    dataset:"sales",

    category:"Math",

    formula:"=SUM(C2:C100)",

    syntax:"SUM(number1,[number2],...)",

    explanation:"Adds all numbers in a range and returns the total value.",

    breakdown:[
        "C2:C100 = Amount Range",
        "Adds all values",
        "Returns Total"
    ],

    output:"125000",

    business:"Used for sales totals, expense reports, salary summaries, and MIS dashboards.",

    tips:[
        "Use AutoSum for quick totals.",
        "Can add multiple ranges."
    ],

    errors:[
        "#VALUE!"
    ],

    alternative:"=SUBTOTAL(9,C2:C100)"
},


// ==========================================================
// SUMIF
// ==========================================================
{
    title:"SUMIF",

    keywords:[
        "sumif",
        "conditional sum",
        "sum with condition",
        "single criteria total",
        "category wise total",
        "employee wise total",
        "department total",
        "sales by category",
        "expense by type",
        "conditional calculation",
        "excel math function",
        "mis reporting",
        "data analysis",
        "automation",
        "advanced excel",
        "business calculation",
        "financial analysis",
        "criteria based total",
        "filtered total",
        "report automation"
    ],

    aliases:[
        "condition total formula",
        "criteria sum",
        "single condition sum"
    ],

    questions:[
        "How to sum based on condition?",
        "How to calculate department total?",
        "How to use SUMIF?",
        "How to total one category?",
        "How to add values with criteria?"
    ],

    dataset:"sales",

    category:"Math",

    formula:'=SUMIF(B:B,"IT",C:C)',

    syntax:"SUMIF(range,criteria,sum_range)",

    explanation:"Adds values that meet a single condition.",

    breakdown:[
        "B:B = Criteria Column",
        "\"IT\" = Condition",
        "C:C = Sum Column"
    ],

    output:"45000",

    business:"Used for department totals, category analysis, and expense tracking.",

    tips:[
        "Use SUMIFS for multiple conditions.",
        "Criteria can be text, numbers, or expressions."
    ],

    errors:[
        "#VALUE!"
    ],

    alternative:'=SUMIFS(C:C,B:B,"IT")'
},


// ==========================================================
// SUMIFS
// ==========================================================
{
    title:"SUMIFS",

    keywords:[
        "sumifs",
        "multiple criteria sum",
        "conditional total",
        "sum with multiple conditions",
        "department month total",
        "employee salary filter",
        "advanced sum",
        "multiple conditions",
        "criteria based calculation",
        "financial analysis",
        "mis summary",
        "excel math function",
        "automation",
        "advanced excel",
        "business reporting",
        "dashboard calculation",
        "data analysis",
        "multi condition total",
        "report preparation",
        "complex calculation"
    ],

    aliases:[
        "multiple condition sum",
        "advanced sum formula",
        "criteria total formula"
    ],

    questions:[
        "How to sum with multiple criteria?",
        "How to calculate monthly department total?",
        "How to use SUMIFS?",
        "How to add filtered values?",
        "How to create MIS summary?"
    ],

    dataset:"employee_salary",

    category:"Math",

    formula:'=SUMIFS(D:D,B:B,"IT",C:C,"Jan")',

    syntax:"SUMIFS(sum_range,criteria_range1,criteria1,...)",

    explanation:"Adds values that satisfy multiple conditions.",

    breakdown:[
        "D:D = Salary Amount",
        "B:B = Department",
        "C:C = Month",
        "Multiple Criteria Applied"
    ],

    output:"85000",

    business:"Used for MIS summaries, financial reports, and management dashboards.",

    tips:[
        "Best function for multi-condition totals.",
        "Works with dates also."
    ],

    errors:[
        "#VALUE!",
        "Incorrect range size"
    ],

    alternative:'=SUMPRODUCT((B:B="IT")*(C:C="Jan")*D:D)'
},


// ==========================================================
// PRODUCT
// ==========================================================
{
    title:"PRODUCT",

    keywords:[
        "product",
        "multiply values",
        "multiplication formula",
        "calculate multiplication",
        "quantity price calculation",
        "cost calculation",
        "total amount",
        "financial calculation",
        "excel math function",
        "number multiplication",
        "invoice calculation",
        "sales calculation",
        "automation",
        "advanced excel",
        "business calculation",
        "data analysis",
        "report preparation",
        "formula calculation",
        "numeric function"
    ],

    aliases:[
        "multiply function",
        "multiplication formula",
        "multiple numbers formula"
    ],

    questions:[
        "How to multiply values?",
        "How to calculate quantity times price?",
        "How to use PRODUCT function?",
        "How to multiply multiple cells?"
    ],

    dataset:"sales",

    category:"Math",

    formula:"=PRODUCT(B2:C2)",

    syntax:"PRODUCT(number1,[number2],...)",

    explanation:"Multiplies all numbers provided as arguments.",

    breakdown:[
        "B2 = Quantity",
        "C2 = Price",
        "Returns Total Amount"
    ],

    output:"5000",

    business:"Used for invoice calculations, sales amount, and financial models.",

    tips:[
        "Useful when multiplying many cells.",
        "For two cells, * operator is simpler."
    ],

    errors:[
        "#VALUE!"
    ],

    alternative:"=B2*C2"
},


// ==========================================================
// ROUND
// ==========================================================
{
    title:"ROUND",

    keywords:[
        "round",
        "round number",
        "decimal rounding",
        "remove decimals",
        "financial rounding",
        "amount formatting",
        "decimal control",
        "invoice rounding",
        "report formatting",
        "number precision",
        "excel math function",
        "calculation formula",
        "automation",
        "advanced excel",
        "financial analysis",
        "accounting formula",
        "numeric formatting",
        "business reporting",
        "data presentation"
    ],

    aliases:[
        "decimal remover",
        "number rounding formula",
        "precision control"
    ],

    questions:[
        "How to round numbers?",
        "How to remove decimals?",
        "How to round amount?",
        "How to use ROUND function?",
        "How to control decimal places?"
    ],

    dataset:"finance",

    category:"Math",

    formula:"=ROUND(A2,2)",

    syntax:"ROUND(number,num_digits)",

    explanation:"Rounds a number to the specified number of digits.",

    breakdown:[
        "A2 = Number",
        "2 = Decimal Places",
        "Returns Rounded Value"
    ],

    output:"1250.75",

    business:"Used in financial reports, invoices, and accounting calculations.",

    tips:[
        "Use ROUNDUP to always increase.",
        "Use ROUNDDOWN to always decrease."
    ],

    errors:[
        "#VALUE!"
    ],

    alternative:"=MROUND(A2,5)"
},

// ==========================================================
// ROUNDUP
// ==========================================================
{
    title:"ROUNDUP",

    keywords:[
        "roundup",
        "round up",
        "always increase rounding",
        "increase decimal value",
        "ceiling rounding",
        "financial rounding",
        "invoice rounding",
        "amount adjustment",
        "decimal control",
        "number precision",
        "excel math function",
        "calculation formula",
        "accounting formula",
        "report formatting",
        "automation",
        "advanced excel",
        "business calculation",
        "numeric function",
        "data presentation",
        "financial analysis"
    ],

    aliases:[
        "increase rounding formula",
        "always higher value",
        "decimal increase"
    ],

    questions:[
        "How to round number upward?",
        "How to always increase decimals?",
        "How to use ROUNDUP?",
        "How to round invoice amount?"
    ],

    dataset:"finance",

    category:"Math",

    formula:"=ROUNDUP(A2,0)",

    syntax:"ROUNDUP(number,num_digits)",

    explanation:"Rounds a number away from zero to the specified number of digits.",

    breakdown:[
        "A2 = Number",
        "0 = No Decimal",
        "Returns Higher Rounded Value"
    ],

    output:"126",

    business:"Used for billing, quantity planning, and financial calculations.",

    tips:[
        "Always rounds upward.",
        "Useful when minimum quantity is required."
    ],

    errors:[
        "#VALUE!"
    ],

    alternative:"=CEILING(A2,1)"
},


// ==========================================================
// ROUNDDOWN
// ==========================================================
{
    title:"ROUNDDOWN",

    keywords:[
        "rounddown",
        "round down",
        "always decrease rounding",
        "lower rounding",
        "floor rounding",
        "decimal removal",
        "financial calculation",
        "amount adjustment",
        "number precision",
        "invoice calculation",
        "excel math function",
        "calculation formula",
        "accounting report",
        "automation",
        "advanced excel",
        "business calculation",
        "numeric function",
        "data formatting",
        "financial analysis"
    ],

    aliases:[
        "decrease rounding formula",
        "always lower value",
        "floor number"
    ],

    questions:[
        "How to round number downward?",
        "How to reduce decimals?",
        "How to use ROUNDDOWN?",
        "How to always get lower value?"
    ],

    dataset:"finance",

    category:"Math",

    formula:"=ROUNDDOWN(A2,0)",

    syntax:"ROUNDDOWN(number,num_digits)",

    explanation:"Rounds a number toward zero.",

    breakdown:[
        "A2 = Number",
        "0 = No Decimal",
        "Returns Lower Rounded Value"
    ],

    output:"125",

    business:"Used for conservative calculations and quantity management.",

    tips:[
        "Does not increase value.",
        "Useful for capacity planning."
    ],

    errors:[
        "#VALUE!"
    ],

    alternative:"=FLOOR(A2,1)"
},


// ==========================================================
// MROUND
// ==========================================================
{
    title:"MROUND",

    keywords:[
        "mround",
        "nearest multiple",
        "round to multiple",
        "multiple rounding",
        "nearest 5",
        "nearest 10",
        "price rounding",
        "quantity rounding",
        "pack size calculation",
        "financial rounding",
        "excel math function",
        "number formatting",
        "automation",
        "advanced excel",
        "business calculation",
        "inventory calculation",
        "sales pricing",
        "report preparation",
        "numeric analysis"
    ],

    aliases:[
        "multiple round formula",
        "nearest number rounding",
        "step rounding"
    ],

    questions:[
        "How to round to nearest 5?",
        "How to round to nearest 10?",
        "How to use MROUND?",
        "How to round prices?"
    ],

    dataset:"inventory",

    category:"Math",

    formula:"=MROUND(A2,5)",

    syntax:"MROUND(number,multiple)",

    explanation:"Rounds a number to the nearest specified multiple.",

    breakdown:[
        "A2 = Number",
        "5 = Multiple Value",
        "Returns Nearest Multiple"
    ],

    output:"125",

    business:"Used for pricing, packaging, and inventory calculations.",

    tips:[
        "Great for standardizing numbers.",
        "Multiple can be any valid number."
    ],

    errors:[
        "#NUM!",
        "#VALUE!"
    ],

    alternative:"=ROUND(A2/5,0)*5"
},


// ==========================================================
// ABS
// ==========================================================
{
    title:"ABS",

    keywords:[
        "abs",
        "absolute value",
        "remove negative sign",
        "positive value",
        "variance calculation",
        "difference analysis",
        "actual vs budget",
        "reconciliation",
        "financial variance",
        "error amount",
        "excel math function",
        "calculation formula",
        "automation",
        "advanced excel",
        "business analysis",
        "reporting",
        "data cleaning",
        "numeric function",
        "comparison analysis",
        "finance report"
    ],

    aliases:[
        "absolute difference",
        "positive number formula",
        "ignore negative sign"
    ],

    questions:[
        "How to remove negative sign?",
        "How to calculate absolute difference?",
        "How to use ABS function?",
        "How to find variance value?"
    ],

    dataset:"finance",

    category:"Math",

    formula:"=ABS(B2-C2)",

    syntax:"ABS(number)",

    explanation:"Returns the absolute value of a number without its sign.",

    breakdown:[
        "B2 = Actual Value",
        "C2 = Target Value",
        "Returns Positive Difference"
    ],

    output:"500",

    business:"Used for variance analysis, reconciliation, and financial comparison.",

    tips:[
        "Useful when direction of difference does not matter.",
        "Common in audit reports."
    ],

    errors:[
        "#VALUE!"
    ],

    alternative:"=IF(B2>C2,B2-C2,C2-B2)"
},


// ==========================================================
// MOD
// ==========================================================
{
    title:"MOD",

    keywords:[
        "mod",
        "remainder",
        "division remainder",
        "leftover value",
        "even odd check",
        "number division",
        "cycle calculation",
        "employee shift rotation",
        "batch calculation",
        "excel math function",
        "calculation formula",
        "automation",
        "advanced excel",
        "business calculation",
        "data analysis",
        "number validation",
        "sequence formula",
        "pattern creation",
        "numeric function"
    ],

    aliases:[
        "remaining value formula",
        "division balance",
        "leftover calculation"
    ],

    questions:[
        "How to find remainder?",
        "How to check even odd number?",
        "How to use MOD function?",
        "How to calculate leftover value?"
    ],

    dataset:"operations",

    category:"Math",

    formula:"=MOD(A2,10)",

    syntax:"MOD(number,divisor)",

    explanation:"Returns the remainder after division.",

    breakdown:[
        "A2 = Number",
        "10 = Divisor",
        "Returns Remaining Value"
    ],

    output:"5",

    business:"Used for rotation logic, batch processing, and validation checks.",

    tips:[
        "MOD(number,2) checks even/odd.",
        "Useful in scheduling formulas."
    ],

    errors:[
        "#DIV/0!"
    ],

    alternative:"=A2-(INT(A2/10)*10)"
},

// ==========================================================
// COUNT
// ==========================================================
{
    title:"COUNT",

    keywords:[
        "count",
        "count numbers",
        "number of records",
        "count numeric values",
        "count entries",
        "total numbers",
        "employee count",
        "sales count",
        "transaction count",
        "data counting",
        "excel counting function",
        "kpi calculation",
        "dashboard metric",
        "mis report",
        "data analysis",
        "automation",
        "advanced excel",
        "record count",
        "numeric data",
        "business reporting"
    ],

    aliases:[
        "number count formula",
        "numeric counter",
        "count values"
    ],

    questions:[
        "How to count numbers?",
        "How to count records?",
        "How to use COUNT function?",
        "How many numeric values exist?",
        "How to count transactions?"
    ],

    priorityRules:[
        "If query contains numeric values => COUNT",
        "If query contains non blank => COUNTA",
        "If query contains blank => COUNTBLANK"
    ],

    dataset:"sales",

    category:"Counting",

    formula:"=COUNT(C2:C100)",

    syntax:"COUNT(value1,[value2],...)",

    explanation:"Counts cells containing numeric values.",

    breakdown:[
        "C2:C100 = Data Range",
        "Counts Only Numbers",
        "Ignores Text and Blank Cells"
    ],

    output:"85",

    business:"Used for KPI calculation, transaction count, and numerical analysis.",

    tips:[
        "COUNT ignores text values.",
        "Use COUNTA for all filled cells."
    ],

    errors:[
        "#VALUE!"
    ],

    alternative:"=COUNTA(C2:C100)"
},


// ==========================================================
// COUNTA
// ==========================================================
{
    title:"COUNTA",

    keywords:[
        "counta",
        "count non blank",
        "count filled cells",
        "count records",
        "count entries",
        "filled cell count",
        "employee records",
        "customer count",
        "data availability",
        "excel counting function",
        "kpi calculation",
        "dashboard metric",
        "mis report",
        "data analysis",
        "automation",
        "advanced excel",
        "record tracking",
        "database count",
        "business reporting",
        "data validation"
    ],

    aliases:[
        "non empty count",
        "filled value counter",
        "record counter"
    ],

    questions:[
        "How to count non blank cells?",
        "How to count records?",
        "How to use COUNTA?",
        "How many cells contain data?",
        "How to count filled entries?"
    ],

    dataset:"employees",

    category:"Counting",

    formula:"=COUNTA(A2:A100)",

    syntax:"COUNTA(value1,[value2],...)",

    explanation:"Counts all non-empty cells including text, numbers, and errors.",

    breakdown:[
        "A2:A100 = Data Range",
        "Counts Filled Cells",
        "Ignores Blank Cells"
    ],

    output:"95",

    business:"Used for record counting, employee database checks, and data completeness.",

    tips:[
        "Counts text and numbers.",
        "Useful before data analysis."
    ],

    errors:[
        "#VALUE!"
    ],

    alternative:"=COUNT(A2:A100)"
},


// ==========================================================
// COUNTBLANK
// ==========================================================
{
    title:"COUNTBLANK",

    keywords:[
        "countblank",
        "blank cells",
        "empty cells",
        "missing data",
        "check blanks",
        "find empty records",
        "data quality check",
        "incomplete data",
        "missing values",
        "excel counting function",
        "data cleaning",
        "audit check",
        "validation report",
        "automation",
        "advanced excel",
        "business reporting",
        "data analysis",
        "quality control",
        "error detection",
        "dashboard"
    ],

    aliases:[
        "empty cell counter",
        "blank finder",
        "missing value counter"
    ],

    questions:[
        "How to count blank cells?",
        "How to find missing data?",
        "How to check empty records?",
        "How to use COUNTBLANK?",
        "How many cells are empty?"
    ],

    dataset:"employee_master",

    category:"Counting",

    formula:"=COUNTBLANK(D2:D100)",

    syntax:"COUNTBLANK(range)",

    explanation:"Counts the number of empty cells in a range.",

    breakdown:[
        "D2:D100 = Data Range",
        "Checks Empty Cells",
        "Returns Blank Count"
    ],

    output:"12",

    business:"Used for data quality checks and incomplete record identification.",

    tips:[
        "Useful before reporting.",
        "Helps identify missing information."
    ],

    errors:[
        "No common errors"
    ],

    alternative:'=COUNTIF(D2:D100,"")'
},


// ==========================================================
// COUNTIF
// ==========================================================
{
    title:"COUNTIF",

    keywords:[
        "countif",
        "count with condition",
        "conditional count",
        "status count",
        "employee status count",
        "category count",
        "single criteria count",
        "count matching values",
        "attendance count",
        "sales category count",
        "excel counting function",
        "kpi calculation",
        "dashboard report",
        "mis reporting",
        "automation",
        "advanced excel",
        "business analysis",
        "data analysis",
        "criteria based count",
        "record analysis"
    ],

    aliases:[
        "condition counter",
        "criteria count formula",
        "filtered count"
    ],

    questions:[
        "How to count based on condition?",
        "How many employees are active?",
        "How to use COUNTIF?",
        "How to count matching values?",
        "How to count category wise?"
    ],

    dataset:"employees",

    category:"Counting",

    formula:'=COUNTIF(C:C,"Active")',

    syntax:"COUNTIF(range,criteria)",

    explanation:"Counts cells that meet a specific condition.",

    breakdown:[
        "C:C = Status Column",
        "\"Active\" = Condition",
        "Counts Matching Records"
    ],

    output:"75",

    business:"Used for status reports, attendance analysis, and KPI dashboards.",

    tips:[
        "Use COUNTIFS for multiple conditions.",
        "Supports wildcards."
    ],

    errors:[
        "#VALUE!"
    ],

    alternative:'=COUNTIFS(C:C,"Active")'
},


// ==========================================================
// COUNTIFS
// ==========================================================
{
    title:"COUNTIFS",

    keywords:[
        "countifs",
        "multiple criteria count",
        "count multiple conditions",
        "advanced count",
        "employee department count",
        "monthly record count",
        "multi condition analysis",
        "conditional counting",
        "complex criteria",
        "dashboard kpi",
        "mis summary",
        "excel counting function",
        "automation",
        "advanced excel",
        "business reporting",
        "data analysis",
        "filtered records",
        "multiple filters",
        "criteria based report",
        "management dashboard"
    ],

    aliases:[
        "multiple condition counter",
        "advanced count formula",
        "criteria based counter"
    ],

    questions:[
        "How to count with multiple conditions?",
        "How to count department wise monthly data?",
        "How to use COUNTIFS?",
        "How to create KPI count?",
        "How to count filtered records?"
    ],

    dataset:"attendance",

    category:"Counting",

    formula:'=COUNTIFS(B:B,"IT",C:C,"Present")',

    syntax:"COUNTIFS(criteria_range1,criteria1,...)",

    explanation:"Counts records that satisfy multiple conditions.",

    breakdown:[
        "B:B = Department",
        "C:C = Attendance Status",
        "Multiple Conditions Applied"
    ],

    output:"45",

    business:"Used for MIS dashboards, HR analytics, and management reporting.",

    tips:[
        "Best for KPI dashboards.",
        "Works with dates and text criteria."
    ],

    errors:[
        "#VALUE!",
        "Incorrect range size"
    ],

    alternative:'=SUMPRODUCT((B:B="IT")*(C:C="Present"))'
},

// ==========================================================
// LEFT
// ==========================================================
{
    title:"LEFT",

    keywords:[
        "left",
        "left characters",
        "extract first characters",
        "get first letters",
        "starting text",
        "text extraction",
        "employee code extraction",
        "customer id extraction",
        "prefix extraction",
        "remove ending text",
        "excel text function",
        "data cleaning",
        "text formatting",
        "automation",
        "advanced excel",
        "master data cleaning",
        "report preparation",
        "string manipulation",
        "text analysis",
        "business reporting"
    ],

    aliases:[
        "first character formula",
        "start text extraction",
        "prefix finder"
    ],

    questions:[
        "How to extract first characters?",
        "How to get first letters from text?",
        "How to use LEFT function?",
        "How to extract employee code prefix?",
        "How to split text from left?"
    ],

    dataset:"employees",

    category:"Text",

    formula:'=LEFT(A2,3)',

    syntax:"LEFT(text,num_chars)",

    explanation:"Returns the specified number of characters from the beginning of a text string.",

    breakdown:[
        "A2 = Text Value",
        "3 = Number of Characters",
        "Returns Starting Characters"
    ],

    output:"EMP",

    business:"Used for employee IDs, product codes, and text extraction.",

    tips:[
        "Useful for fixed-length codes.",
        "Combine with LEN for dynamic extraction."
    ],

    errors:[
        "#VALUE!"
    ],

    alternative:"=TEXTBEFORE(A2,\"-\")"
},


// ==========================================================
// RIGHT
// ==========================================================
{
    title:"RIGHT",

    keywords:[
        "right",
        "right characters",
        "extract last characters",
        "get ending text",
        "last digits",
        "suffix extraction",
        "account number last digits",
        "employee code extraction",
        "text extraction",
        "excel text function",
        "data cleaning",
        "text formatting",
        "automation",
        "advanced excel",
        "master data",
        "report preparation",
        "string manipulation",
        "text analysis",
        "business reporting"
    ],

    aliases:[
        "last character formula",
        "ending text extraction",
        "suffix finder"
    ],

    questions:[
        "How to extract last characters?",
        "How to get last digits?",
        "How to use RIGHT function?",
        "How to extract suffix?",
        "How to split text from right?"
    ],

    dataset:"customers",

    category:"Text",

    formula:'=RIGHT(A2,4)',

    syntax:"RIGHT(text,num_chars)",

    explanation:"Returns the specified number of characters from the end of a text string.",

    breakdown:[
        "A2 = Text Value",
        "4 = Number of Characters",
        "Returns Ending Characters"
    ],

    output:"4589",

    business:"Used for account numbers, codes, and identifier extraction.",

    tips:[
        "Useful for extracting last digits.",
        "Combine with LEN for dynamic formulas."
    ],

    errors:[
        "#VALUE!"
    ],

    alternative:"=TEXTAFTER(A2,\"-\")"
},


// ==========================================================
// MID
// ==========================================================
{
    title:"MID",

    keywords:[
        "mid",
        "middle characters",
        "extract middle text",
        "text between positions",
        "substring formula",
        "employee id extraction",
        "code extraction",
        "text parsing",
        "string extraction",
        "excel text function",
        "data cleaning",
        "automation",
        "advanced excel",
        "master data",
        "report preparation",
        "text analysis",
        "business reporting",
        "text manipulation",
        "data transformation",
        "string handling"
    ],

    aliases:[
        "middle text formula",
        "substring function",
        "extract characters"
    ],

    questions:[
        "How to extract middle characters?",
        "How to use MID function?",
        "How to extract text from position?",
        "How to split code values?",
        "How to get characters between text?"
    ],

    dataset:"employee_codes",

    category:"Text",

    formula:'=MID(A2,4,3)',

    syntax:"MID(text,start_num,num_chars)",

    explanation:"Returns characters from the middle of a text string based on starting position.",

    breakdown:[
        "A2 = Text Value",
        "4 = Starting Position",
        "3 = Number of Characters"
    ],

    output:"123",

    business:"Used for extracting codes, IDs, and structured text data.",

    tips:[
        "Useful for fixed-format codes.",
        "Combine with FIND for dynamic extraction."
    ],

    errors:[
        "#VALUE!"
    ],

    alternative:'=TEXTBEFORE(TEXTAFTER(A2,"-"),"-")'
},


// ==========================================================
// LEN
// ==========================================================
{
    title:"LEN",

    keywords:[
        "len",
        "text length",
        "count characters",
        "number of characters",
        "character count",
        "data validation",
        "check text length",
        "password length",
        "employee id validation",
        "excel text function",
        "data cleaning",
        "automation",
        "advanced excel",
        "quality check",
        "text analysis",
        "string calculation",
        "report preparation",
        "business validation",
        "master data",
        "text formatting"
    ],

    aliases:[
        "character counter",
        "text size formula",
        "length checker"
    ],

    questions:[
        "How to count characters?",
        "How to find text length?",
        "How to use LEN function?",
        "How to validate ID length?",
        "How many characters are there?"
    ],

    dataset:"employee_master",

    category:"Text",

    formula:"=LEN(A2)",

    syntax:"LEN(text)",

    explanation:"Returns the number of characters in a text string including spaces.",

    breakdown:[
        "A2 = Text Value",
        "Counts Characters",
        "Returns Length"
    ],

    output:"8",

    business:"Used for data validation, code checking, and text analysis.",

    tips:[
        "Combine with TRIM to remove unwanted spaces.",
        "Useful for quality checks."
    ],

    errors:[
        "#VALUE!"
    ],

    alternative:"=LEN(TRIM(A2))"
},


// ==========================================================
// TRIM
// ==========================================================
{
    title:"TRIM",

    keywords:[
        "trim",
        "remove spaces",
        "extra spaces",
        "clean text",
        "remove unwanted spaces",
        "data cleaning",
        "duplicate matching",
        "master data cleanup",
        "imported data cleaning",
        "excel text function",
        "text formatting",
        "automation",
        "advanced excel",
        "data quality",
        "report preparation",
        "text correction",
        "string cleanup",
        "business reporting",
        "data transformation",
        "database cleanup"
    ],

    aliases:[
        "space remover",
        "text cleaner",
        "clean spaces formula"
    ],

    questions:[
        "How to remove extra spaces?",
        "How to clean imported data?",
        "How to use TRIM function?",
        "How to fix text spacing?",
        "How to clean names?"
    ],

    dataset:"customer_master",

    category:"Text",

    formula:"=TRIM(A2)",

    syntax:"TRIM(text)",

    explanation:"Removes extra spaces from text except single spaces between words.",

    breakdown:[
        "A2 = Text Value",
        "Removes Leading Spaces",
        "Removes Extra Spaces"
    ],

    output:"John Smith",

    business:"Used for cleaning imported data and improving matching accuracy.",

    tips:[
        "Use before VLOOKUP/XLOOKUP matching.",
        "Very useful in Power Query cleanup."
    ],

    errors:[
        "#VALUE!"
    ],

    alternative:"=CLEAN(TRIM(A2))"
},

// ==========================================================
// CONCAT
// ==========================================================
{
    title:"CONCAT",

    keywords:[
        "concat",
        "combine text",
        "join text",
        "merge cells",
        "combine columns",
        "combine names",
        "create full name",
        "merge values",
        "text combination",
        "employee name creation",
        "email creation",
        "excel text function",
        "data formatting",
        "automation",
        "advanced excel",
        "report preparation",
        "text manipulation",
        "business reporting",
        "data transformation",
        "string joining"
    ],

    aliases:[
        "text combine formula",
        "join cells formula",
        "merge text function"
    ],

    questions:[
        "How to combine text?",
        "How to merge first and last name?",
        "How to use CONCAT function?",
        "How to join multiple cells?",
        "How to create full name?"
    ],

    dataset:"employees",

    category:"Text",

    formula:"=CONCAT(A2,B2)",

    syntax:"CONCAT(text1,[text2],...)",

    explanation:"Combines text from multiple cells into one text string.",

    breakdown:[
        "A2 = First Name",
        "B2 = Last Name",
        "Returns Combined Text"
    ],

    output:"JohnSmith",

    business:"Used for creating names, IDs, and combined fields.",

    tips:[
        "Use TEXTJOIN when you need separators.",
        "Modern replacement for CONCATENATE."
    ],

    errors:[
        "#VALUE!"
    ],

    alternative:'=TEXTJOIN(" ",TRUE,A2,B2)'
},


// ==========================================================
// CONCATENATE
// ==========================================================
{
    title:"CONCATENATE",

    keywords:[
        "concatenate",
        "old concatenate formula",
        "combine text",
        "join cells",
        "merge values",
        "text joining",
        "combine columns",
        "create full name",
        "old excel text function",
        "legacy formula",
        "excel text function",
        "data formatting",
        "automation",
        "advanced excel",
        "report preparation",
        "string manipulation",
        "text analysis",
        "business reporting",
        "data transformation"
    ],

    aliases:[
        "old concat formula",
        "join text formula",
        "merge string formula"
    ],

    questions:[
        "How to use CONCATENATE?",
        "How to combine text in old Excel?",
        "How to join two cells?",
        "How to merge values?"
    ],

    dataset:"employees",

    category:"Text",

    formula:'=CONCATENATE(A2," ",B2)',

    syntax:"CONCATENATE(text1,text2,...)",

    explanation:"Joins multiple text values into one string.",

    breakdown:[
        "A2 = First Name",
        "\" \" = Space",
        "B2 = Last Name"
    ],

    output:"John Smith",

    business:"Used in older Excel versions for combining text.",

    tips:[
        "CONCAT is preferred in new Excel.",
        "TEXTJOIN is better for separators."
    ],

    errors:[
        "#VALUE!"
    ],

    alternative:'=CONCAT(A2," ",B2)'
},


// ==========================================================
// TEXTJOIN
// ==========================================================
{
    title:"TEXTJOIN",

    keywords:[
        "textjoin",
        "join text with separator",
        "combine with comma",
        "merge multiple cells",
        "combine rows",
        "combine columns",
        "create list",
        "comma separated values",
        "dynamic text combination",
        "employee list",
        "email list",
        "excel text function",
        "automation",
        "advanced excel",
        "report preparation",
        "data formatting",
        "text manipulation",
        "business reporting",
        "data transformation"
    ],

    aliases:[
        "delimiter join formula",
        "comma join formula",
        "advanced concat"
    ],

    questions:[
        "How to combine text with comma?",
        "How to merge multiple cells?",
        "How to use TEXTJOIN?",
        "How to create comma separated list?",
        "How to ignore blank cells?"
    ],

    dataset:"employees",

    category:"Text",

    formula:'=TEXTJOIN(", ",TRUE,A2:A10)',

    syntax:"TEXTJOIN(delimiter,ignore_empty,text1,...)",

    explanation:"Combines multiple text values using a chosen separator.",

    breakdown:[
        "\", \" = Separator",
        "TRUE = Ignore Blank Cells",
        "A2:A10 = Text Range"
    ],

    output:"John, David, Peter",

    business:"Used for email lists, reports, and dynamic summaries.",

    tips:[
        "Best option for combining many cells.",
        "Supports ranges."
    ],

    errors:[
        "#VALUE!"
    ],

    alternative:"=CONCAT(A2:A10)"
},


// ==========================================================
// FIND
// ==========================================================
{
    title:"FIND",

    keywords:[
        "find",
        "case sensitive search",
        "find text position",
        "locate character",
        "text position",
        "search within text",
        "character location",
        "string analysis",
        "text extraction",
        "data parsing",
        "excel text function",
        "automation",
        "advanced excel",
        "text cleaning",
        "report preparation",
        "data transformation",
        "business analysis",
        "text validation",
        "pattern finding"
    ],

    aliases:[
        "character finder",
        "position finder",
        "text location formula"
    ],

    questions:[
        "How to find text position?",
        "How to locate character?",
        "How to use FIND function?",
        "How to check if text exists?",
        "How to extract based on position?"
    ],

    dataset:"customer_data",

    category:"Text",

    formula:'=FIND("-",A2)',

    syntax:"FIND(find_text,within_text,[start_num])",

    explanation:"Returns the position of a text string inside another text string. It is case-sensitive.",

    breakdown:[
        "\"-\" = Search Character",
        "A2 = Text Source",
        "Returns Character Position"
    ],

    output:"5",

    business:"Used for text parsing and extracting structured information.",

    tips:[
        "Case-sensitive function.",
        "Use SEARCH for case-insensitive search."
    ],

    errors:[
        "#VALUE!",
        "Text not found"
    ],

    alternative:'=SEARCH("-",A2)'
},


// ==========================================================
// SEARCH
// ==========================================================
{
    title:"SEARCH",

    keywords:[
        "search",
        "case insensitive search",
        "find text",
        "locate word",
        "text search",
        "keyword search",
        "contains text",
        "find phrase",
        "data classification",
        "text analysis",
        "excel text function",
        "automation",
        "advanced excel",
        "report preparation",
        "business analysis",
        "text filtering",
        "pattern matching",
        "string search",
        "data validation"
    ],

    aliases:[
        "contains formula",
        "text finder",
        "keyword finder"
    ],

    questions:[
        "How to search text inside cell?",
        "How to check if word exists?",
        "How to use SEARCH function?",
        "How to find keyword?",
        "How to locate text position?"
    ],

    dataset:"customer_feedback",

    category:"Text",

    formula:'=SEARCH("urgent",A2)',

    syntax:"SEARCH(find_text,within_text,[start_num])",

    explanation:"Finds the position of text inside another text string without case sensitivity.",

    breakdown:[
        "\"urgent\" = Keyword",
        "A2 = Text Source",
        "Returns Position"
    ],

    output:"8",

    business:"Used for keyword analysis, feedback classification, and text filtering.",

    tips:[
        "Not case-sensitive.",
        "Use with ISNUMBER for text detection."
    ],

    errors:[
        "#VALUE!",
        "Text not found"
    ],

    alternative:'=ISNUMBER(SEARCH("urgent",A2))'
},

// ==========================================================
// SUBSTITUTE
// ==========================================================
{
    title:"SUBSTITUTE",

    keywords:[
        "substitute",
        "replace word",
        "replace text",
        "change text",
        "replace specific value",
        "remove character",
        "text correction",
        "data cleaning",
        "replace old value",
        "fix spelling",
        "excel text function",
        "automation",
        "advanced excel",
        "master data cleanup",
        "text formatting",
        "report preparation",
        "string manipulation",
        "business reporting",
        "data transformation",
        "text cleanup"
    ],

    aliases:[
        "text replacement formula",
        "word replace function",
        "value replacement"
    ],

    questions:[
        "How to replace text?",
        "How to remove unwanted characters?",
        "How to use SUBSTITUTE?",
        "How to change old values?",
        "How to replace specific word?"
    ],

    dataset:"customer_master",

    category:"Text",

    formula:'=SUBSTITUTE(A2,"Mumbai","Pune")',

    syntax:"SUBSTITUTE(text,old_text,new_text,[instance_num])",

    explanation:"Replaces existing text with new text in a string.",

    breakdown:[
        "A2 = Original Text",
        "Mumbai = Old Value",
        "Pune = New Value"
    ],

    output:"Pune",

    business:"Used for data cleanup, correction, and standardization.",

    tips:[
        "Can replace specific occurrence using instance_num.",
        "Useful in imported data cleaning."
    ],

    errors:[
        "#VALUE!"
    ],

    alternative:'=REPLACE(A2,1,6,"Pune")'
},


// ==========================================================
// REPLACE
// ==========================================================
{
    title:"REPLACE",

    keywords:[
        "replace",
        "replace by position",
        "change characters",
        "overwrite text",
        "modify string",
        "replace fixed length text",
        "text editing",
        "data correction",
        "mask account number",
        "hide sensitive data",
        "excel text function",
        "automation",
        "advanced excel",
        "data cleaning",
        "text formatting",
        "report preparation",
        "string manipulation",
        "business reporting",
        "data security"
    ],

    aliases:[
        "position replace formula",
        "character replacement",
        "text overwrite"
    ],

    questions:[
        "How to replace characters by position?",
        "How to mask account number?",
        "How to use REPLACE function?",
        "How to change fixed text length?"
    ],

    dataset:"bank_accounts",

    category:"Text",

    formula:'=REPLACE(A2,5,4,"XXXX")',

    syntax:"REPLACE(old_text,start_num,num_chars,new_text)",

    explanation:"Replaces characters in a text string based on position.",

    breakdown:[
        "A2 = Original Text",
        "5 = Starting Position",
        "4 = Characters To Replace",
        "XXXX = New Text"
    ],

    output:"1234XXXX7890",

    business:"Used for masking sensitive data and modifying fixed-format text.",

    tips:[
        "Use SUBSTITUTE when matching actual text.",
        "Use REPLACE when position matters."
    ],

    errors:[
        "#VALUE!"
    ],

    alternative:'=SUBSTITUTE(A2,"5678","XXXX")'
},


// ==========================================================
// CLEAN
// ==========================================================
{
    title:"CLEAN",

    keywords:[
        "clean",
        "remove hidden characters",
        "remove non printable characters",
        "data cleaning",
        "import cleanup",
        "remove unwanted symbols",
        "fix copied data",
        "excel text cleaning",
        "remove line break",
        "dirty data",
        "master data cleanup",
        "excel text function",
        "automation",
        "advanced excel",
        "data quality",
        "report preparation",
        "text correction",
        "business reporting",
        "data transformation"
    ],

    aliases:[
        "text cleaner",
        "hidden character remover",
        "import data cleaner"
    ],

    questions:[
        "How to remove hidden characters?",
        "How to clean copied data?",
        "How to use CLEAN function?",
        "How to remove unwanted symbols?"
    ],

    dataset:"imported_data",

    category:"Text",

    formula:"=CLEAN(A2)",

    syntax:"CLEAN(text)",

    explanation:"Removes non-printable characters from text.",

    breakdown:[
        "A2 = Dirty Text",
        "Removes Hidden Characters",
        "Returns Clean Text"
    ],

    output:"Employee Name",

    business:"Used for Power Query preparation and imported file cleanup.",

    tips:[
        "Combine with TRIM for complete cleaning.",
        "Useful after copying data from websites."
    ],

    errors:[
        "#VALUE!"
    ],

    alternative:"=TRIM(CLEAN(A2))"
},


// ==========================================================
// EXACT
// ==========================================================
{
    title:"EXACT",

    keywords:[
        "exact",
        "compare text",
        "exact text comparison",
        "case sensitive comparison",
        "match two values",
        "check duplicate",
        "validate text",
        "compare names",
        "data verification",
        "quality check",
        "excel text function",
        "automation",
        "advanced excel",
        "data validation",
        "audit check",
        "master data",
        "report preparation",
        "business analysis",
        "text matching"
    ],

    aliases:[
        "text comparison formula",
        "case sensitive match",
        "string comparison"
    ],

    questions:[
        "How to compare two texts?",
        "How to check exact match?",
        "How to use EXACT function?",
        "How to find duplicate text?"
    ],

    dataset:"employee_master",

    category:"Text",

    formula:"=EXACT(A2,B2)",

    syntax:"EXACT(text1,text2)",

    explanation:"Checks whether two text values are exactly identical including case.",

    breakdown:[
        "A2 = First Text",
        "B2 = Second Text",
        "Returns TRUE/FALSE"
    ],

    output:"TRUE",

    business:"Used for audit checks, duplicate validation, and data comparison.",

    tips:[
        "Case-sensitive comparison.",
        "Use = for normal comparison."
    ],

    errors:[
        "#VALUE!"
    ],

    alternative:'=A2=B2'
},


// ==========================================================
// TEXT
// ==========================================================
{
    title:"TEXT",

    keywords:[
        "text",
        "format number",
        "convert number to text",
        "date formatting",
        "currency formatting",
        "percentage formatting",
        "custom format",
        "display format",
        "excel text function",
        "number formatting",
        "report preparation",
        "dashboard formatting",
        "automation",
        "advanced excel",
        "financial report",
        "date conversion",
        "business reporting",
        "presentation format",
        "data formatting"
    ],

    aliases:[
        "number format formula",
        "date format formula",
        "custom display formula"
    ],

    questions:[
        "How to format date?",
        "How to convert number to text?",
        "How to display currency format?",
        "How to use TEXT function?"
    ],

    dataset:"finance",

    category:"Text",

    formula:'=TEXT(A2,"$#,##0")',

    syntax:"TEXT(value,format_text)",

    explanation:"Converts a value into text with a specified format.",

    breakdown:[
        "A2 = Value",
        "$#,##0 = Format Pattern",
        "Returns Formatted Text"
    ],

    output:"$65,000",

    business:"Used for reports, dashboards, invoices, and presentation formatting.",

    tips:[
        "TEXT changes value into text.",
        "Use only for display purposes."
    ],

    errors:[
        "#VALUE!"
    ],

    alternative:'=TEXT(A2,"dd-mmm-yyyy")'
},

// ==========================================================
// TODAY
// ==========================================================
{
    title:"TODAY",

    keywords:[
        "today",
        "today date",
        "current date",
        "today formula",
        "system date",
        "automatic date",
        "date today",
        "dynamic date",
        "daily report date",
        "mis date",
        "dashboard date",
        "excel date function",
        "date tracking",
        "automation",
        "advanced excel",
        "report preparation",
        "compliance date",
        "due date calculation",
        "business reporting",
        "date analysis"
    ],

    aliases:[
        "current day formula",
        "auto date formula",
        "live date"
    ],

    questions:[
        "How to get today's date?",
        "How to insert current date automatically?",
        "How to use TODAY function?",
        "How to create dynamic date?"
    ],

    dataset:"reporting",

    category:"Date",

    formula:"=TODAY()",

    syntax:"TODAY()",

    explanation:"Returns the current system date automatically.",

    breakdown:[
        "No arguments required",
        "Uses system date",
        "Updates automatically"
    ],

    output:"28-Jul-2026",

    business:"Used in MIS reports, aging analysis, and deadline tracking.",

    tips:[
        "Updates when workbook recalculates.",
        "Use VALUE paste if you need a fixed date."
    ],

    errors:[
        "No common errors"
    ],

    alternative:"=NOW()"
},


// ==========================================================
// NOW
// ==========================================================
{
    title:"NOW",

    keywords:[
        "now",
        "current date time",
        "current timestamp",
        "date and time",
        "live time",
        "system timestamp",
        "time tracking",
        "audit timestamp",
        "transaction time",
        "report generation time",
        "excel date function",
        "automation",
        "advanced excel",
        "dashboard timestamp",
        "business reporting",
        "activity tracking",
        "log creation",
        "monitoring",
        "data analysis",
        "workflow tracking"
    ],

    aliases:[
        "timestamp formula",
        "live date time",
        "current time formula"
    ],

    questions:[
        "How to get current date and time?",
        "How to create timestamp?",
        "How to use NOW function?",
        "How to track update time?"
    ],

    dataset:"transactions",

    category:"Date",

    formula:"=NOW()",

    syntax:"NOW()",

    explanation:"Returns the current date and time from the system.",

    breakdown:[
        "No arguments required",
        "Returns Date",
        "Returns Current Time"
    ],

    output:"28-Jul-2026 08:30 PM",

    business:"Used for audit logs, timestamps, and automated reporting.",

    tips:[
        "Updates automatically.",
        "Useful for tracking refresh time."
    ],

    errors:[
        "No common errors"
    ],

    alternative:"=TODAY()"
},


// ==========================================================
// DATE
// ==========================================================
{
    title:"DATE",

    keywords:[
        "date",
        "create date",
        "combine year month day",
        "generate date",
        "date creation",
        "manual date formula",
        "convert numbers to date",
        "due date formula",
        "joining date",
        "expiry date",
        "excel date function",
        "automation",
        "advanced excel",
        "report preparation",
        "compliance tracking",
        "business reporting",
        "date calculation",
        "calendar formula",
        "data analysis",
        "period creation"
    ],

    aliases:[
        "date builder",
        "date generator",
        "year month day formula"
    ],

    questions:[
        "How to create date from numbers?",
        "How to combine year month day?",
        "How to use DATE function?",
        "How to generate custom dates?"
    ],

    dataset:"employee_master",

    category:"Date",

    formula:"=DATE(A2,B2,C2)",

    syntax:"DATE(year,month,day)",

    explanation:"Creates a valid Excel date from separate year, month, and day values.",

    breakdown:[
        "A2 = Year",
        "B2 = Month",
        "C2 = Day"
    ],

    output:"15-Jan-2026",

    business:"Used for joining dates, expiry dates, and compliance schedules.",

    tips:[
        "Useful when date components are stored separately.",
        "Excel automatically handles month/year overflow."
    ],

    errors:[
        "#VALUE!",
        "#NUM!"
    ],

    alternative:"=TEXT(DATE(A2,B2,C2),\"dd-mmm-yyyy\")"
},


// ==========================================================
// YEAR
// ==========================================================
{
    title:"YEAR",

    keywords:[
        "year",
        "extract year",
        "get year from date",
        "date year",
        "financial year",
        "year calculation",
        "joining year",
        "expiry year",
        "annual report",
        "year wise analysis",
        "excel date function",
        "automation",
        "advanced excel",
        "mis reporting",
        "dashboard analysis",
        "date extraction",
        "business reporting",
        "trend analysis",
        "period analysis"
    ],

    aliases:[
        "date year formula",
        "year extraction",
        "annual grouping"
    ],

    questions:[
        "How to extract year from date?",
        "How to get financial year?",
        "How to use YEAR function?",
        "How to group data by year?"
    ],

    dataset:"sales",

    category:"Date",

    formula:"=YEAR(A2)",

    syntax:"YEAR(serial_number)",

    explanation:"Returns the year component from a date.",

    breakdown:[
        "A2 = Date Value",
        "Extracts Year",
        "Returns Four Digit Year"
    ],

    output:"2026",

    business:"Used for yearly MIS, trend analysis, and reporting.",

    tips:[
        "Works with valid Excel dates only.",
        "Useful with Pivot reports."
    ],

    errors:[
        "#VALUE!"
    ],

    alternative:'=TEXT(A2,"YYYY")'
},


// ==========================================================
// MONTH
// ==========================================================
{
    title:"MONTH",

    keywords:[
        "month",
        "extract month",
        "get month from date",
        "month number",
        "monthly report",
        "month wise analysis",
        "period calculation",
        "financial month",
        "sales month",
        "attendance month",
        "excel date function",
        "automation",
        "advanced excel",
        "mis reporting",
        "dashboard analysis",
        "date extraction",
        "business reporting",
        "trend analysis",
        "calendar analysis",
        "period grouping"
    ],

    aliases:[
        "date month formula",
        "month extraction",
        "monthly grouping"
    ],

    questions:[
        "How to extract month from date?",
        "How to get month number?",
        "How to use MONTH function?",
        "How to create monthly reports?"
    ],

    dataset:"attendance",

    category:"Date",

    formula:"=MONTH(A2)",

    syntax:"MONTH(serial_number)",

    explanation:"Returns the month number from a date.",

    breakdown:[
        "A2 = Date Value",
        "Extracts Month",
        "Returns 1 to 12"
    ],

    output:"7",

    business:"Used for monthly MIS, attendance analysis, and trend reports.",

    tips:[
        "Combine with TEXT for month names.",
        "Useful in Pivot grouping."
    ],

    errors:[
        "#VALUE!"
    ],

    alternative:'=TEXT(A2,"MMMM")'
},

// ==========================================================
// DAY
// ==========================================================
{
    title:"DAY",

    keywords:[
        "day",
        "extract day",
        "get day from date",
        "day number",
        "date day",
        "calendar day",
        "joining day",
        "transaction day",
        "attendance day",
        "invoice date day",
        "excel date function",
        "automation",
        "advanced excel",
        "mis reporting",
        "dashboard analysis",
        "date extraction",
        "business reporting",
        "calendar analysis",
        "period calculation",
        "date validation"
    ],

    aliases:[
        "day extraction",
        "date day formula",
        "day number formula"
    ],

    questions:[
        "How to extract day from date?",
        "How to get day number?",
        "How to use DAY function?",
        "How to analyze daily data?"
    ],

    dataset:"attendance",

    category:"Date",

    formula:"=DAY(A2)",

    syntax:"DAY(serial_number)",

    explanation:"Returns the day number from a date.",

    breakdown:[
        "A2 = Date Value",
        "Extracts Day",
        "Returns 1 to 31"
    ],

    output:"28",

    business:"Used for daily reports, attendance tracking, and date analysis.",

    tips:[
        "Works only with valid Excel dates.",
        "Combine with MONTH and YEAR for reporting."
    ],

    errors:[
        "#VALUE!"
    ],

    alternative:'=TEXT(A2,"DD")'
},


// ==========================================================
// DATEDIF
// ==========================================================
{
    title:"DATEDIF",

    keywords:[
        "datedif",
        "date difference",
        "age calculation",
        "calculate age",
        "employee experience",
        "years between dates",
        "months between dates",
        "days between dates",
        "tenure calculation",
        "service period",
        "joining duration",
        "retirement calculation",
        "excel date function",
        "automation",
        "advanced excel",
        "hr analytics",
        "employee report",
        "business reporting",
        "date analysis",
        "duration formula"
    ],

    aliases:[
        "age formula",
        "experience formula",
        "tenure formula"
    ],

    questions:[
        "How to calculate age?",
        "How to calculate experience?",
        "How many years between dates?",
        "How to use DATEDIF?",
        "How to calculate employee tenure?"
    ],

    priorityRules:[
        "If query contains age => prefer DATEDIF",
        "If query contains date difference days => prefer DAYS"
    ],

    dataset:"employees",

    category:"Date",

    formula:'=DATEDIF(B2,TODAY(),"Y")',

    syntax:"DATEDIF(start_date,end_date,unit)",

    explanation:"Calculates the difference between two dates in years, months, or days.",

    breakdown:[
        "B2 = Joining/Birth Date",
        "TODAY() = Current Date",
        "\"Y\" = Complete Years"
    ],

    output:"8 Years",

    business:"Used for employee experience, age calculation, and tenure reports.",

    tips:[
        "Hidden Excel function but widely used.",
        "Units: Y, M, D."
    ],

    errors:[
        "#NUM!",
        "#VALUE!"
    ],

    alternative:'=YEARFRAC(B2,TODAY())'
},


// ==========================================================
// DAYS
// ==========================================================
{
    title:"DAYS",

    keywords:[
        "days",
        "date difference",
        "number of days",
        "days between dates",
        "remaining days",
        "deadline calculation",
        "due date tracking",
        "leave days",
        "project duration",
        "turnaround time",
        "sla days",
        "excel date function",
        "automation",
        "advanced excel",
        "mis reporting",
        "business reporting",
        "date calculation",
        "period analysis",
        "compliance tracking",
        "workflow monitoring"
    ],

    aliases:[
        "days difference formula",
        "date gap formula",
        "duration days"
    ],

    questions:[
        "How to calculate days between dates?",
        "How many days remaining?",
        "How to calculate SLA days?",
        "How to use DAYS function?"
    ],

    dataset:"compliance",

    category:"Date",

    formula:"=DAYS(B2,A2)",

    syntax:"DAYS(end_date,start_date)",

    explanation:"Returns the number of days between two dates.",

    breakdown:[
        "B2 = End Date",
        "A2 = Start Date",
        "Returns Day Difference"
    ],

    output:"30",

    business:"Used for SLA tracking, leave calculation, and compliance monitoring.",

    tips:[
        "Use TODAY() for pending days calculation.",
        "Works with actual Excel dates."
    ],

    errors:[
        "#VALUE!"
    ],

    alternative:"=B2-A2"
},


// ==========================================================
// EDATE
// ==========================================================
{
    title:"EDATE",

    keywords:[
        "edate",
        "add months",
        "subtract months",
        "future date",
        "previous date",
        "renewal date",
        "expiry date",
        "maturity date",
        "contract period",
        "subscription end date",
        "loan calculation",
        "excel date function",
        "automation",
        "advanced excel",
        "business reporting",
        "date calculation",
        "compliance tracking",
        "schedule management",
        "period calculation",
        "financial analysis"
    ],

    aliases:[
        "month addition formula",
        "date shift formula",
        "future month date"
    ],

    questions:[
        "How to add months to date?",
        "How to calculate expiry date?",
        "How to move date forward?",
        "How to use EDATE function?"
    ],

    dataset:"contracts",

    category:"Date",

    formula:"=EDATE(A2,12)",

    syntax:"EDATE(start_date,months)",

    explanation:"Returns a date after adding or subtracting specified months.",

    breakdown:[
        "A2 = Start Date",
        "12 = Number of Months",
        "Returns Future Date"
    ],

    output:"28-Jul-2027",

    business:"Used for renewal dates, contracts, and financial schedules.",

    tips:[
        "Negative months move backward.",
        "Useful for monthly planning."
    ],

    errors:[
        "#VALUE!"
    ],

    alternative:"=DATE(YEAR(A2)+1,MONTH(A2),DAY(A2))"
},


// ==========================================================
// EOMONTH
// ==========================================================
{
    title:"EOMONTH",

    keywords:[
        "eomonth",
        "month end",
        "last date of month",
        "month closing date",
        "financial month end",
        "period end",
        "report closing",
        "month end report",
        "accounting date",
        "salary processing date",
        "month boundary",
        "excel date function",
        "automation",
        "advanced excel",
        "mis reporting",
        "financial reporting",
        "business reporting",
        "date calculation",
        "period analysis",
        "calendar formula"
    ],

    aliases:[
        "month end formula",
        "last day formula",
        "closing date formula"
    ],

    questions:[
        "How to find month end date?",
        "How to calculate last date of month?",
        "How to use EOMONTH?",
        "How to create month closing report?"
    ],

    dataset:"finance",

    category:"Date",

    formula:"=EOMONTH(A2,0)",

    syntax:"EOMONTH(start_date,months)",

    explanation:"Returns the last day of a month before or after a specified date.",

    breakdown:[
        "A2 = Starting Date",
        "0 = Current Month",
        "Returns Month End Date"
    ],

    output:"31-Jul-2026",

    business:"Used for month-end closing, MIS reports, and financial reporting.",

    tips:[
        "Use -1 for previous month end.",
        "Use 1 for next month end."
    ],

    errors:[
        "#VALUE!"
    ],

    alternative:"=DATE(YEAR(A2),MONTH(A2)+1,0)"
},

//
// ==========================================================
// VLOOKUP
// ==========================================================
{
    title:"VLOOKUP",

    keywords:[
        "vlookup",
        "vertical lookup",
        "find employee",
        "employee salary",
        "lookup value",
        "search table",
        "retrieve data",
        "fetch value",
        "match id",
        "lookup formula",
        "find record",
        "database lookup",
        "excel lookup function",
        "table search",
        "master data lookup",
        "employee details",
        "salary lookup",
        "customer lookup",
        "old lookup formula",
        "excel reference function"
    ],

    aliases:[
        "vertical search",
        "table lookup",
        "column lookup"
    ],

    priorityRules:[
        "If query contains vertical lookup => VLOOKUP",
        "If query contains modern lookup => XLOOKUP",
        "If query contains lookup only => show VLOOKUP and XLOOKUP"
    ],

    questions:[
        "How to find employee salary?",
        "How to fetch value from another table?",
        "How to use VLOOKUP?",
        "How to lookup employee details?"
    ],

    dataset:"employees",

    category:"Lookup",

    formula:"=VLOOKUP(H2,A2:E100,4,FALSE)",

    syntax:"VLOOKUP(value,table,col_index,FALSE)",

    explanation:"Searches the first column of a table and returns a matching value from another column.",

    breakdown:[
        "H2 = Employee ID",
        "A:E = Employee Table",
        "4 = Salary Column",
        "FALSE = Exact Match"
    ],

    output:"65000",

    business:"Used for employee lookup, salary reports, customer master data, and MIS automation.",

    tips:[
        "Lookup column must be first column.",
        "Use XLOOKUP for modern Excel."
    ],

    errors:[
        "#N/A",
        "#REF!",
        "Wrong column index"
    ],

    alternative:"=XLOOKUP(H2,A:A,D:D)"
},


// ==========================================================
// XLOOKUP
// ==========================================================
{
    title:"XLOOKUP",

    keywords:[
        "xlookup",
        "modern lookup",
        "advanced lookup",
        "find value",
        "search employee",
        "fetch salary",
        "lookup data",
        "retrieve information",
        "match record",
        "dynamic lookup",
        "replace vlookup",
        "new lookup formula",
        "excel lookup function",
        "table search",
        "database lookup",
        "employee details",
        "customer details",
        "flexible lookup",
        "exact match lookup",
        "excel reference function"
    ],

    aliases:[
        "new vlookup",
        "advanced search formula",
        "modern excel lookup"
    ],

    priorityRules:[
        "If query contains xlookup => XLOOKUP",
        "If query contains modern lookup => XLOOKUP",
        "If query contains lookup only => show XLOOKUP"
    ],

    questions:[
        "How to replace VLOOKUP?",
        "How to use XLOOKUP?",
        "How to find matching value?",
        "How to lookup from another table?"
    ],

    dataset:"employees",

    category:"Lookup",

    formula:"=XLOOKUP(H2,A2:A100,D2:D100,\"Not Found\")",

    syntax:"XLOOKUP(lookup_value,lookup_array,return_array,[if_not_found])",

    explanation:"Searches a range and returns the corresponding value from another range.",

    breakdown:[
        "H2 = Search Value",
        "A:A = Lookup Range",
        "D:D = Return Range",
        "Not Found = Error Handling"
    ],

    output:"65000",

    business:"Used for modern dashboards, automated reports, and flexible data retrieval.",

    tips:[
        "Can lookup left or right.",
        "Includes built-in error handling."
    ],

    errors:[
        "#N/A",
        "#VALUE!"
    ],

    alternative:"=INDEX(D:D,MATCH(H2,A:A,0))"
},


// ==========================================================
// HLOOKUP
// ==========================================================
{
    title:"HLOOKUP",

    keywords:[
        "hlookup",
        "horizontal lookup",
        "row lookup",
        "search horizontally",
        "lookup across row",
        "monthly data lookup",
        "header lookup",
        "table row search",
        "retrieve horizontal value",
        "excel lookup function",
        "reference formula",
        "dashboard lookup",
        "report automation",
        "data retrieval",
        "business reporting",
        "excel advanced",
        "table search",
        "matrix lookup",
        "horizontal search"
    ],

    aliases:[
        "row lookup",
        "horizontal search formula"
    ],

    priorityRules:[
        "If query contains horizontal lookup => HLOOKUP"
    ],

    questions:[
        "How to lookup horizontally?",
        "How to search values in rows?",
        "How to use HLOOKUP?"
    ],

    dataset:"monthly_sales",

    category:"Lookup",

    formula:"=HLOOKUP(B1,A1:M5,4,FALSE)",

    syntax:"HLOOKUP(value,table,row_index,FALSE)",

    explanation:"Searches the first row of a table and returns a value from a specified row.",

    breakdown:[
        "B1 = Month",
        "A1:M5 = Table Range",
        "4 = Return Row"
    ],

    output:"125000",

    business:"Used for monthly reports and horizontal data structures.",

    tips:[
        "Use XLOOKUP for new Excel versions.",
        "Useful for row-based tables."
    ],

    errors:[
        "#N/A",
        "#REF!"
    ],

    alternative:"=XLOOKUP(B1,A1:M1,A4:M4)"
},


// ==========================================================
// INDEX
// ==========================================================
{
    title:"INDEX",

    keywords:[
        "index",
        "return value by position",
        "find value position",
        "array lookup",
        "dynamic lookup",
        "two way lookup",
        "retrieve value",
        "table reference",
        "advanced lookup",
        "matrix formula",
        "excel lookup function",
        "reference function",
        "data retrieval",
        "dashboard formula",
        "automation",
        "advanced excel",
        "business reporting",
        "database search",
        "position based lookup"
    ],

    aliases:[
        "position lookup",
        "array return formula",
        "value extraction formula"
    ],

    priorityRules:[
        "If query contains two way lookup => INDEX + MATCH"
    ],

    questions:[
        "How to return value by position?",
        "How to use INDEX?",
        "How to create two way lookup?"
    ],

    dataset:"sales",

    category:"Lookup",

    formula:"=INDEX(D2:D100,MATCH(H2,A2:A100,0))",

    syntax:"INDEX(array,row_num,[column_num])",

    explanation:"Returns a value from a table based on row and column position.",

    breakdown:[
        "D2:D100 = Return Range",
        "MATCH = Finds Position",
        "INDEX = Returns Value"
    ],

    output:"65000",

    business:"Used for advanced lookup models and dynamic reports.",

    tips:[
        "Combine with MATCH for flexible lookup.",
        "More powerful than traditional lookup."
    ],

    errors:[
        "#REF!",
        "#VALUE!"
    ],

    alternative:"=XLOOKUP(H2,A:A,D:D)"
},


// ==========================================================
// MATCH
// ==========================================================
{
    title:"MATCH",

    keywords:[
        "match",
        "find position",
        "lookup position",
        "search position",
        "row number",
        "column number",
        "find index",
        "locate value",
        "position formula",
        "dynamic lookup",
        "excel lookup function",
        "reference function",
        "data analysis",
        "automation",
        "advanced excel",
        "dashboard formula",
        "table search",
        "value matching",
        "lookup helper",
        "position finder"
    ],

    aliases:[
        "position finder",
        "lookup index",
        "location formula"
    ],

    priorityRules:[
        "If query contains match position => MATCH",
        "If query contains exact lookup => combine with INDEX"
    ],

    questions:[
        "How to find position of value?",
        "How to use MATCH?",
        "How to find row number?",
        "How to create dynamic lookup?"
    ],

    dataset:"employees",

    category:"Lookup",

    formula:"=MATCH(H2,A2:A100,0)",

    syntax:"MATCH(value,array,match_type)",

    explanation:"Returns the position of a value in a range.",

    breakdown:[
        "H2 = Search Value",
        "A2:A100 = Search Range",
        "0 = Exact Match"
    ],

    output:"15",

    business:"Used with INDEX for advanced lookup and dynamic reporting.",

    tips:[
        "Use 0 for exact matching.",
        "Commonly paired with INDEX."
    ],

    errors:[
        "#N/A"
    ],

    alternative:"=XMATCH(H2,A2:A100)"
},

// ==========================================================
// COUNT
// ==========================================================
{
    title:"COUNT",

    keywords:[
        "count",
        "count numbers",
        "number count",
        "count records",
        "count transactions",
        "count employees",
        "numeric count",
        "data counting",
        "volume analysis",
        "record analysis",
        "excel statistical function",
        "data analysis",
        "mis reporting",
        "dashboard calculation",
        "business reporting",
        "statistics formula",
        "kpi calculation",
        "performance analysis",
        "report automation",
        "data summary"
    ],

    aliases:[
        "number counter",
        "numeric records count",
        "count formula"
    ],

    priorityRules:[
        "If query contains blank cells => COUNTBLANK",
        "If query contains condition => COUNTIF/COUNTIFS",
        "If query contains numbers only => COUNT"
    ],

    questions:[
        "How to count numbers?",
        "How to count records?",
        "How to count employees?"
    ],

    dataset:"employees",

    category:"Statistical",

    formula:"=COUNT(D2:D100)",

    syntax:"COUNT(value1,[value2],...)",

    explanation:"Counts the number of cells containing numeric values.",

    breakdown:[
        "D2:D100 = Data Range",
        "Counts Numeric Values",
        "Ignores Text and Blank Cells"
    ],

    output:"95",

    business:"Used for transaction volume, employee count, and numeric analysis.",

    tips:[
        "Use COUNTA for all non-empty cells.",
        "Use COUNTIF for conditional counting."
    ],

    errors:[
        "#VALUE!"
    ],

    alternative:"=COUNTA(D2:D100)"
},


// ==========================================================
// COUNTA
// ==========================================================
{
    title:"COUNTA",

    keywords:[
        "counta",
        "count filled cells",
        "count non blank",
        "count records",
        "count names",
        "count employee records",
        "filled cell count",
        "data availability",
        "record count",
        "population count",
        "excel statistical function",
        "data analysis",
        "mis reporting",
        "dashboard formula",
        "business reporting",
        "data quality",
        "data validation",
        "statistics formula",
        "report automation",
        "summary calculation"
    ],

    aliases:[
        "non blank count",
        "filled value counter",
        "record counter"
    ],

    priorityRules:[
        "If query contains filled cells => COUNTA",
        "If query contains numbers only => COUNT"
    ],

    questions:[
        "How to count filled cells?",
        "How many records are available?",
        "How to use COUNTA?"
    ],

    dataset:"customer_master",

    category:"Statistical",

    formula:"=COUNTA(A2:A100)",

    syntax:"COUNTA(value1,[value2],...)",

    explanation:"Counts all non-empty cells including text, numbers, and errors.",

    breakdown:[
        "A2:A100 = Data Range",
        "Counts Filled Cells",
        "Ignores Empty Cells"
    ],

    output:"98",

    business:"Used for record count, data completeness, and quality checks.",

    tips:[
        "Best for counting rows with data.",
        "Includes text values."
    ],

    errors:[
        "#VALUE!"
    ],

    alternative:"=ROWS(A2:A100)-COUNTBLANK(A2:A100)"
},

// ==========================================================
// IF
// ==========================================================
{
    title:"IF",

    keywords:[
        "if",
        "if formula",
        "condition formula",
        "logical test",
        "check condition",
        "true false formula",
        "decision formula",
        "status formula",
        "approval formula",
        "pass fail formula",
        "employee status",
        "salary condition",
        "excel logical function",
        "automation",
        "advanced excel",
        "report preparation",
        "business logic",
        "data validation",
        "decision making",
        "conditional calculation"
    ],

    aliases:[
        "conditional formula",
        "if then formula",
        "logic formula"
    ],

    priorityRules:[
        "If exact query contains if => IF",
        "If query contains multiple conditions => IFS",
        "If query contains all conditions => AND"
    ],

    questions:[
        "How to use IF formula?",
        "How to create condition?",
        "How to return value based on condition?"
    ],

    dataset:"employees",

    category:"Logical",

    formula:'=IF(D2>=50000,"High","Low")',

    syntax:"IF(logical_test,value_if_true,value_if_false)",

    explanation:"Checks a condition and returns one value if TRUE and another value if FALSE.",

    breakdown:[
        "D2>=50000 = Condition",
        "\"High\" = If True",
        "\"Low\" = If False"
    ],

    output:"High",

    business:"Used for status checks, approvals, grading, and automated decisions.",

    tips:[
        "Use IFS for multiple conditions.",
        "Combine with AND/OR for complex logic."
    ],

    errors:[
        "#VALUE!",
        "#NAME?"
    ],

    alternative:'=IFS(D2>=75000,"High",D2>=50000,"Medium",TRUE,"Low")'
},


// ==========================================================
// IFS
// ==========================================================
{
    title:"IFS",

    keywords:[
        "ifs",
        "multiple conditions",
        "multiple if",
        "nested if replacement",
        "many conditions",
        "grade calculation",
        "rating formula",
        "salary band",
        "performance rating",
        "condition list",
        "advanced if formula",
        "excel logical function",
        "automation",
        "advanced excel",
        "business rules",
        "decision formula",
        "classification formula",
        "report automation",
        "data analysis",
        "conditional logic"
    ],

    aliases:[
        "multiple if formula",
        "nested if alternative",
        "condition chain formula"
    ],

    priorityRules:[
        "If query contains multiple conditions => IFS",
        "If query contains simple if => IF"
    ],

    questions:[
        "How to handle multiple conditions?",
        "How to replace nested IF?",
        "How to use IFS formula?"
    ],

    dataset:"performance",

    category:"Logical",

    formula:'=IFS(B2>=90,"Excellent",B2>=75,"Good",B2>=50,"Average",TRUE,"Poor")',

    syntax:"IFS(test1,result1,test2,result2,...)",

    explanation:"Checks multiple conditions and returns the result for the first TRUE condition.",

    breakdown:[
        "B2>=90 = First Condition",
        "B2>=75 = Second Condition",
        "TRUE = Default Result"
    ],

    output:"Good",

    business:"Used for grading, ratings, salary bands, and classifications.",

    tips:[
        "Cleaner alternative to nested IF.",
        "Keep conditions in logical order."
    ],

    errors:[
        "#N/A",
        "#VALUE!"
    ],

    alternative:'=IF(B2>=90,"Excellent",IF(B2>=75,"Good","Poor"))'
},


// ==========================================================
// AND
// ==========================================================
{
    title:"AND",

    keywords:[
        "and",
        "all conditions true",
        "multiple checks",
        "combined conditions",
        "logical and",
        "approval criteria",
        "validation formula",
        "eligibility check",
        "employee eligibility",
        "compliance check",
        "excel logical function",
        "automation",
        "advanced excel",
        "business rules",
        "decision making",
        "data validation",
        "audit check",
        "conditional logic",
        "report automation"
    ],

    aliases:[
        "all criteria formula",
        "multiple true check",
        "combined test"
    ],

    priorityRules:[
        "If query contains all conditions => AND",
        "If query contains either condition => OR"
    ],

    questions:[
        "How to check multiple conditions?",
        "How to verify all conditions are true?",
        "How to use AND formula?"
    ],

    dataset:"employees",

    category:"Logical",

    formula:'=AND(B2>=18,C2="Active")',

    syntax:"AND(logical1,[logical2],...)",

    explanation:"Returns TRUE only when all conditions are TRUE.",

    breakdown:[
        "B2>=18 = Age Check",
        "C2=\"Active\" = Status Check",
        "Both Must Be TRUE"
    ],

    output:"TRUE",

    business:"Used for eligibility, approval, and compliance checks.",

    tips:[
        "Usually combined with IF.",
        "All conditions must pass."
    ],

    errors:[
        "#VALUE!"
    ],

    alternative:'=IF(AND(B2>=18,C2="Active"),"Eligible","Not Eligible")'
},


// ==========================================================
// OR
// ==========================================================
{
    title:"OR",

    keywords:[
        "or",
        "any condition true",
        "either condition",
        "multiple option check",
        "logical or",
        "alternative criteria",
        "exception formula",
        "approval condition",
        "validation logic",
        "risk check",
        "excel logical function",
        "automation",
        "advanced excel",
        "business rules",
        "decision formula",
        "data validation",
        "conditional logic",
        "report automation"
    ],

    aliases:[
        "either criteria formula",
        "any match formula",
        "alternative condition"
    ],

    priorityRules:[
        "If query contains any condition => OR",
        "If query contains all conditions => AND"
    ],

    questions:[
        "How to check either condition?",
        "How to use OR formula?",
        "How to test multiple options?"
    ],

    dataset:"transactions",

    category:"Logical",

    formula:'=OR(C2="High",D2="Urgent")',

    syntax:"OR(logical1,[logical2],...)",

    explanation:"Returns TRUE if at least one condition is TRUE.",

    breakdown:[
        "C2=\"High\" = First Check",
        "D2=\"Urgent\" = Second Check",
        "Any TRUE Returns TRUE"
    ],

    output:"TRUE",

    business:"Used for exception reports, alerts, and risk monitoring.",

    tips:[
        "Combine with IF for decisions.",
        "Useful for multiple acceptance criteria."
    ],

    errors:[
        "#VALUE!"
    ],

    alternative:'=IF(OR(C2="High",D2="Urgent"),"Alert","Normal")'
},


// ==========================================================
// NOT
// ==========================================================
{
    title:"NOT",

    keywords:[
        "not",
        "reverse condition",
        "opposite logic",
        "negative condition",
        "exclude condition",
        "not equal check",
        "validation formula",
        "exception logic",
        "logical opposite",
        "excel logical function",
        "automation",
        "advanced excel",
        "business rules",
        "data validation",
        "conditional logic",
        "audit check",
        "report automation",
        "decision formula"
    ],

    aliases:[
        "reverse logic",
        "negative test",
        "opposite formula"
    ],

    priorityRules:[
        "If query contains not equal => NOT"
    ],

    questions:[
        "How to reverse condition?",
        "How to use NOT formula?",
        "How to check opposite value?"
    ],

    dataset:"employees",

    category:"Logical",

    formula:'=NOT(C2="Inactive")',

    syntax:"NOT(logical)",

    explanation:"Reverses the logical value. TRUE becomes FALSE and FALSE becomes TRUE.",

    breakdown:[
        "C2=\"Inactive\" = Condition",
        "NOT Reverses Result",
        "Returns TRUE/FALSE"
    ],

    output:"TRUE",

    business:"Used for exclusions, validations, and exception checks.",

    tips:[
        "Useful with IF.",
        "Alternative to <> operator in some cases."
    ],

    errors:[
        "#VALUE!"
    ],

    alternative:'=C2<>"Inactive"'
},

// ==========================================================
// SUM
// ==========================================================
{
    title:"SUM",

    keywords:[
        "sum",
        "total",
        "add values",
        "calculate total",
        "grand total",
        "sales total",
        "salary total",
        "expense total",
        "amount calculation",
        "numeric addition",
        "excel math function",
        "calculation formula",
        "mis reporting",
        "dashboard calculation",
        "financial analysis",
        "business reporting",
        "data summary",
        "automation",
        "advanced excel",
        "aggregation"
    ],

    aliases:[
        "addition formula",
        "total formula",
        "sum calculation"
    ],

    priorityRules:[
        "If query contains condition => SUMIF",
        "If query contains multiple conditions => SUMIFS",
        "If query contains only total => SUM"
    ],

    questions:[
        "How to calculate total?",
        "How to add numbers?",
        "How to use SUM formula?"
    ],

    dataset:"sales",

    category:"Math",

    formula:"=SUM(D2:D100)",

    syntax:"SUM(number1,[number2],...)",

    explanation:"Adds all numbers in a selected range.",

    breakdown:[
        "D2:D100 = Amount Range",
        "Adds All Values",
        "Returns Total"
    ],

    output:"1250000",

    business:"Used for sales totals, expenses, salaries, and MIS summaries.",

    tips:[
        "Supports multiple ranges.",
        "Ignores text values."
    ],

    errors:[
        "#VALUE!"
    ],

    alternative:"=SUM(D2:D50,D60:D100)"
},


// ==========================================================
// SUMIF
// ==========================================================
{
    title:"SUMIF",

    keywords:[
        "sumif",
        "sum with condition",
        "conditional sum",
        "total by criteria",
        "department wise total",
        "region wise sales",
        "category total",
        "single condition sum",
        "criteria based total",
        "conditional calculation",
        "excel math function",
        "mis reporting",
        "dashboard formula",
        "financial analysis",
        "business reporting",
        "data analysis",
        "automation",
        "advanced excel",
        "smart reporting",
        "aggregation"
    ],

    aliases:[
        "criteria sum",
        "conditional total",
        "single criteria addition"
    ],

    priorityRules:[
        "If query contains one condition => SUMIF",
        "If query contains multiple conditions => SUMIFS"
    ],

    questions:[
        "How to sum based on condition?",
        "How to calculate department total?",
        "How to use SUMIF?"
    ],

    dataset:"employees",

    category:"Math",

    formula:'=SUMIF(C2:C100,"IT",D2:D100)',

    syntax:"SUMIF(range,criteria,sum_range)",

    explanation:"Adds values that meet one specific condition.",

    breakdown:[
        "C2:C100 = Criteria Range",
        "\"IT\" = Condition",
        "D2:D100 = Sum Range"
    ],

    output:"450000",

    business:"Used for department totals, category analysis, and cost reports.",

    tips:[
        "Use SUMIFS for multiple conditions.",
        "Supports wildcards."
    ],

    errors:[
        "#VALUE!"
    ],

    alternative:'=SUMIFS(D:D,C:C,"IT")'
},


// ==========================================================
// SUMIFS
// ==========================================================
{
    title:"SUMIFS",

    keywords:[
        "sumifs",
        "multiple condition sum",
        "sum with multiple criteria",
        "advanced sum",
        "conditional total multiple",
        "department location total",
        "region product sales",
        "complex calculation",
        "multi criteria formula",
        "advanced reporting",
        "excel math function",
        "mis reporting",
        "dashboard calculation",
        "financial analysis",
        "business reporting",
        "data analysis",
        "automation",
        "advanced excel",
        "smart reporting",
        "aggregation"
    ],

    aliases:[
        "multiple criteria sum",
        "advanced conditional sum",
        "complex total formula"
    ],

    priorityRules:[
        "If query contains multiple conditions => SUMIFS"
    ],

    questions:[
        "How to sum with multiple conditions?",
        "How to calculate region and product total?",
        "How to use SUMIFS?"
    ],

    dataset:"sales",

    category:"Math",

    formula:'=SUMIFS(E2:E100,B2:B100,"North",C2:C100,"Laptop")',

    syntax:"SUMIFS(sum_range,criteria_range1,criteria1,...)",

    explanation:"Adds values that satisfy multiple conditions.",

    breakdown:[
        "E2:E100 = Sales Amount",
        "B Column = Region",
        "C Column = Product"
    ],

    output:"850000",

    business:"Used for advanced MIS, financial reports, and KPI analysis.",

    tips:[
        "Most used formula for MIS reporting.",
        "Supports multiple criteria."
    ],

    errors:[
        "#VALUE!"
    ],

    alternative:'=SUMIF(B:B,"North",E:E)'
},


// ==========================================================
// ROUND
// ==========================================================
{
    title:"ROUND",

    keywords:[
        "round",
        "round number",
        "decimal rounding",
        "remove decimals",
        "financial rounding",
        "currency rounding",
        "percentage rounding",
        "report formatting",
        "number precision",
        "decimal control",
        "excel math function",
        "calculation formula",
        "financial analysis",
        "dashboard formatting",
        "mis reporting",
        "automation",
        "advanced excel",
        "data presentation",
        "business reporting"
    ],

    aliases:[
        "decimal round formula",
        "number formatting formula",
        "precision formula"
    ],

    priorityRules:[
        "If query contains always round up => ROUNDUP",
        "If query contains always round down => ROUNDDOWN"
    ],

    questions:[
        "How to round numbers?",
        "How to remove decimal values?",
        "How to use ROUND formula?"
    ],

    dataset:"finance",

    category:"Math",

    formula:"=ROUND(A2,2)",

    syntax:"ROUND(number,num_digits)",

    explanation:"Rounds a number to a specified number of digits.",

    breakdown:[
        "A2 = Number",
        "2 = Decimal Places",
        "Returns Rounded Value"
    ],

    output:"1250.50",

    business:"Used in financial reports, calculations, and presentations.",

    tips:[
        "ROUND changes the actual value.",
        "Formatting only changes display."
    ],

    errors:[
        "#VALUE!"
    ],

    alternative:"=MROUND(A2,5)"
},


// ==========================================================
// ROUNDUP
// ==========================================================
{
    title:"ROUNDUP",

    keywords:[
        "roundup",
        "always round up",
        "increase decimal",
        "ceiling rounding",
        "billing calculation",
        "quantity rounding",
        "minimum unit calculation",
        "financial calculation",
        "price rounding",
        "number adjustment",
        "excel math function",
        "calculation formula",
        "business reporting",
        "automation",
        "advanced excel",
        "data analysis",
        "financial analysis",
        "report preparation",
        "number formatting"
    ],

    aliases:[
        "upward rounding",
        "ceiling number formula",
        "next number rounding"
    ],

    priorityRules:[
        "If query contains always round up => ROUNDUP",
        "If query contains nearest rounding => ROUND"
    ],

    questions:[
        "How to round value upward?",
        "How to always increase decimal?",
        "How to use ROUNDUP?"
    ],

    dataset:"billing",

    category:"Math",

    formula:"=ROUNDUP(A2,0)",

    syntax:"ROUNDUP(number,num_digits)",

    explanation:"Rounds a number away from zero.",

    breakdown:[
        "A2 = Number",
        "0 = No Decimal",
        "Returns Higher Value"
    ],

    output:"1251",

    business:"Used for billing, quantity planning, and minimum requirement calculations.",

    tips:[
        "Always moves value upward.",
        "Useful for capacity planning."
    ],

    errors:[
        "#VALUE!"
    ],

    alternative:"=CEILING(A2,1)"
},


// ==========================================================
// COUNTBLANK
// ==========================================================
{
    title:"COUNTBLANK",

    keywords:[
        "countblank",
        "count blank cells",
        "empty cells",
        "missing data",
        "blank records",
        "data gap",
        "incomplete data",
        "missing values",
        "quality check",
        "data validation",
        "excel statistical function",
        "data analysis",
        "mis reporting",
        "dashboard monitoring",
        "business reporting",
        "data cleaning",
        "audit check",
        "exception reporting",
        "statistics formula",
        "data quality"
    ],

    aliases:[
        "empty cell counter",
        "missing data count",
        "blank checker"
    ],

    priorityRules:[
        "If query contains blank or empty => COUNTBLANK"
    ],

    questions:[
        "How to count blank cells?",
        "How to find missing data?",
        "How to check empty records?"
    ],

    dataset:"employee_master",

    category:"Statistical",

    formula:"=COUNTBLANK(A2:A100)",

    syntax:"COUNTBLANK(range)",

    explanation:"Counts the number of empty cells in a range.",

    breakdown:[
        "A2:A100 = Data Range",
        "Checks Empty Cells",
        "Returns Blank Count"
    ],

    output:"5",

    business:"Used for data quality checks and incomplete record monitoring.",

    tips:[
        "Useful before reporting.",
        "Combine with COUNTA for completeness percentage."
    ],

    errors:[
        "#VALUE!"
    ],

    alternative:"=ROWS(A2:A100)-COUNTA(A2:A100)"
},


// ==========================================================
// COUNTIF
// ==========================================================
{
    title:"COUNTIF",

    keywords:[
        "countif",
        "count with condition",
        "conditional count",
        "count employees by department",
        "count sales by region",
        "criteria count",
        "single condition count",
        "status count",
        "error count",
        "attendance count",
        "excel statistical function",
        "data analysis",
        "mis reporting",
        "dashboard calculation",
        "business reporting",
        "kpi analysis",
        "performance tracking",
        "statistics formula",
        "smart reporting",
        "conditional formula"
    ],

    aliases:[
        "criteria counter",
        "condition count",
        "single criteria count"
    ],

    priorityRules:[
        "If query contains one condition => COUNTIF",
        "If query contains multiple conditions => COUNTIFS"
    ],

    questions:[
        "How to count based on condition?",
        "How many employees in IT department?",
        "How to use COUNTIF?"
    ],

    dataset:"employees",

    category:"Statistical",

    formula:'=COUNTIF(C2:C100,"IT")',

    syntax:"COUNTIF(range,criteria)",

    explanation:"Counts cells that meet a specified condition.",

    breakdown:[
        "C2:C100 = Department Range",
        "\"IT\" = Condition",
        "Returns Matching Count"
    ],

    output:"25",

    business:"Used for department count, status tracking, and KPI reports.",

    tips:[
        "Supports wildcards.",
        "Use COUNTIFS for multiple criteria."
    ],

    errors:[
        "#VALUE!"
    ],

    alternative:'=COUNTIFS(C:C,"IT")'
},


// ==========================================================
// COUNTIFS
// ==========================================================
{
    title:"COUNTIFS",

    keywords:[
        "countifs",
        "multiple condition count",
        "count with multiple criteria",
        "advanced count",
        "conditional counting",
        "department location count",
        "multi criteria count",
        "complex count formula",
        "attendance analysis",
        "error analysis",
        "transaction count",
        "excel statistical function",
        "data analysis",
        "mis reporting",
        "dashboard calculation",
        "business reporting",
        "kpi tracking",
        "statistics formula",
        "advanced excel",
        "smart reporting"
    ],

    aliases:[
        "multiple criteria counter",
        "advanced count formula",
        "multi condition count"
    ],

    priorityRules:[
        "If query contains multiple conditions => COUNTIFS"
    ],

    questions:[
        "How to count with multiple conditions?",
        "How many employees by department and location?",
        "How to use COUNTIFS?"
    ],

    dataset:"employees",

    category:"Statistical",

    formula:'=COUNTIFS(C2:C100,"IT",D2:D100,"Active")',

    syntax:"COUNTIFS(criteria_range1,criteria1,...)",

    explanation:"Counts records that satisfy multiple conditions.",

    breakdown:[
        "C2:C100 = Department",
        "IT = First Condition",
        "Active = Second Condition"
    ],

    output:"18",

    business:"Used for advanced MIS reports, audit checks, and operational analysis.",

    tips:[
        "Supports unlimited criteria pairs.",
        "Ideal for dashboard metrics."
    ],

    errors:[
        "#VALUE!"
    ],

    alternative:'=COUNTIF(C:C,"IT")'
},

// ==========================================================
// ROUNDDOWN
// ==========================================================
{
    title:"ROUNDDOWN",

    keywords:[
        "rounddown",
        "round down",
        "always round down",
        "decrease decimal",
        "truncate number",
        "lower rounding",
        "decimal reduction",
        "financial rounding",
        "price calculation",
        "quantity calculation",
        "excel math function",
        "calculation formula",
        "number adjustment",
        "report formatting",
        "financial analysis",
        "automation",
        "advanced excel",
        "business reporting",
        "data calculation",
        "precision control"
    ],

    aliases:[
        "downward rounding",
        "lower value rounding",
        "truncate formula"
    ],

    priorityRules:[
        "If query contains round down => ROUNDDOWN",
        "If query contains round up => ROUNDUP"
    ],

    questions:[
        "How to round down value?",
        "How to remove decimal downward?",
        "How to use ROUNDDOWN?"
    ],

    dataset:"finance",

    category:"Math",

    formula:"=ROUNDDOWN(A2,0)",

    syntax:"ROUNDDOWN(number,num_digits)",

    explanation:"Rounds a number downward toward zero.",

    breakdown:[
        "A2 = Number",
        "0 = No Decimal",
        "Returns Lower Value"
    ],

    output:"1250",

    business:"Used for billing, quantity calculation, and financial adjustments.",

    tips:[
        "Always rounds downward.",
        "Useful where overestimation is not allowed."
    ],

    errors:[
        "#VALUE!"
    ],

    alternative:"=INT(A2)"
},


// ==========================================================
// CEILING
// ==========================================================
{
    title:"CEILING",

    keywords:[
        "ceiling",
        "round up multiple",
        "nearest multiple",
        "increase to multiple",
        "capacity calculation",
        "package quantity",
        "batch calculation",
        "price rounding",
        "minimum requirement",
        "financial rounding",
        "excel math function",
        "calculation formula",
        "planning calculation",
        "inventory analysis",
        "business reporting",
        "automation",
        "advanced excel",
        "number adjustment",
        "quantity planning"
    ],

    aliases:[
        "multiple rounding",
        "next multiple formula",
        "upper rounding"
    ],

    priorityRules:[
        "If query contains nearest multiple upward => CEILING",
        "If query contains lower multiple => FLOOR"
    ],

    questions:[
        "How to round to nearest multiple?",
        "How to increase value to next multiple?",
        "How to use CEILING?"
    ],

    dataset:"inventory",

    category:"Math",

    formula:"=CEILING(A2,10)",

    syntax:"CEILING(number,significance)",

    explanation:"Rounds a number up to the nearest specified multiple.",

    breakdown:[
        "A2 = Number",
        "10 = Multiple",
        "Returns Next Multiple"
    ],

    output:"130",

    business:"Used for packaging, capacity planning, and quantity calculations.",

    tips:[
        "Useful for business rules requiring minimum units.",
        "Use FLOOR for downward multiples."
    ],

    errors:[
        "#NUM!",
        "#VALUE!"
    ],

    alternative:"=ROUNDUP(A2/10,0)*10"
},


// ==========================================================
// FLOOR
// ==========================================================
{
    title:"FLOOR",

    keywords:[
        "floor",
        "round down multiple",
        "nearest lower multiple",
        "reduce to multiple",
        "quantity rounding",
        "inventory calculation",
        "batch size",
        "financial rounding",
        "price adjustment",
        "minimum threshold",
        "excel math function",
        "calculation formula",
        "planning analysis",
        "business reporting",
        "automation",
        "advanced excel",
        "number adjustment",
        "data calculation",
        "precision control"
    ],

    aliases:[
        "lower multiple formula",
        "downward multiple rounding",
        "minimum multiple"
    ],

    priorityRules:[
        "If query contains nearest lower multiple => FLOOR",
        "If query contains nearest upper multiple => CEILING"
    ],

    questions:[
        "How to round down to multiple?",
        "How to get lower multiple?",
        "How to use FLOOR?"
    ],

    dataset:"inventory",

    category:"Math",

    formula:"=FLOOR(A2,10)",

    syntax:"FLOOR(number,significance)",

    explanation:"Rounds a number down to the nearest specified multiple.",

    breakdown:[
        "A2 = Number",
        "10 = Multiple",
        "Returns Lower Multiple"
    ],

    output:"120",

    business:"Used for inventory units, pricing rules, and batch calculations.",

    tips:[
        "Useful when excess quantity is not required.",
        "Opposite of CEILING."
    ],

    errors:[
        "#NUM!",
        "#VALUE!"
    ],

    alternative:"=ROUNDDOWN(A2/10,0)*10"
},


// ==========================================================
// ABS
// ==========================================================
{
    title:"ABS",

    keywords:[
        "abs",
        "absolute value",
        "remove negative",
        "positive difference",
        "variance calculation",
        "difference without sign",
        "reconciliation difference",
        "financial variance",
        "gap calculation",
        "error amount",
        "excel math function",
        "calculation formula",
        "data analysis",
        "mis reporting",
        "dashboard calculation",
        "financial analysis",
        "audit calculation",
        "automation",
        "advanced excel"
    ],

    aliases:[
        "absolute difference",
        "positive number formula",
        "ignore negative sign"
    ],

    questions:[
        "How to remove negative sign?",
        "How to calculate absolute difference?",
        "How to use ABS formula?"
    ],

    dataset:"reconciliation",

    category:"Math",

    formula:"=ABS(B2-C2)",

    syntax:"ABS(number)",

    explanation:"Returns the absolute value of a number without its sign.",

    breakdown:[
        "B2-C2 = Difference",
        "ABS Removes Negative Sign",
        "Returns Positive Value"
    ],

    output:"500",

    business:"Used for reconciliation, variance analysis, and error measurement.",

    tips:[
        "Commonly used in financial reports.",
        "Useful for comparing actual vs budget."
    ],

    errors:[
        "#VALUE!"
    ],

    alternative:"=IF(B2>C2,B2-C2,C2-B2)"
},


// ==========================================================
// MOD
// ==========================================================
{
    title:"MOD",

    keywords:[
        "mod",
        "remainder",
        "division remainder",
        "remaining value",
        "odd even check",
        "cycle calculation",
        "repeat pattern",
        "allocation calculation",
        "employee rotation",
        "batch calculation",
        "excel math function",
        "calculation formula",
        "automation",
        "advanced excel",
        "data analysis",
        "number calculation",
        "business reporting",
        "planning formula",
        "logic calculation",
        "mathematical function"
    ],

    aliases:[
        "remainder formula",
        "division balance",
        "leftover calculation"
    ],

    priorityRules:[
        "If query contains remainder => MOD",
        "If query contains odd even => MOD"
    ],

    questions:[
        "How to find remainder?",
        "How to check odd or even?",
        "How to use MOD formula?"
    ],

    dataset:"employee_rotation",

    category:"Math",

    formula:"=MOD(A2,7)",

    syntax:"MOD(number,divisor)",

    explanation:"Returns the remainder after division.",

    breakdown:[
        "A2 = Number",
        "7 = Divisor",
        "Returns Remaining Value"
    ],

    output:"3",

    business:"Used for scheduling, rotation logic, and repeating calculations.",

    tips:[
        "Useful with ROW() for repeating patterns.",
        "Can identify odd/even numbers."
    ],

    errors:[
        "#DIV/0!",
        "#VALUE!"
    ],

    alternative:"=A2-(INT(A2/7)*7)"
},

// ==========================================================
// LEFT
// ==========================================================
{
    title:"LEFT",

    keywords:[
        "left",
        "extract first characters",
        "first characters",
        "get starting text",
        "extract beginning",
        "employee code extraction",
        "customer code",
        "prefix extraction",
        "text extraction",
        "split text",
        "excel text function",
        "data cleaning",
        "data preparation",
        "power query cleaning",
        "automation",
        "advanced excel",
        "mis reporting",
        "text analysis",
        "string extraction",
        "text formula"
    ],

    aliases:[
        "start text formula",
        "first letters formula",
        "prefix formula"
    ],

    priorityRules:[
        "If query contains first characters => LEFT",
        "If query contains last characters => RIGHT"
    ],

    questions:[
        "How to extract first characters?",
        "How to get text from beginning?",
        "How to use LEFT formula?"
    ],

    dataset:"employee_codes",

    category:"Text",

    formula:'=LEFT(A2,3)',

    syntax:"LEFT(text,[num_chars])",

    explanation:"Returns characters from the beginning of a text string.",

    breakdown:[
        "A2 = Text Value",
        "3 = Number of Characters",
        "Returns Starting Characters"
    ],

    output:"EMP",

    business:"Used for code extraction, prefixes, and data cleaning.",

    tips:[
        "Useful for fixed-length codes.",
        "Combine with LEN for dynamic extraction."
    ],

    errors:[
        "#VALUE!"
    ],

    alternative:'=TEXTBEFORE(A2,"-")'
},


// ==========================================================
// RIGHT
// ==========================================================
{
    title:"RIGHT",

    keywords:[
        "right",
        "extract last characters",
        "last characters",
        "get ending text",
        "extract end",
        "suffix extraction",
        "employee id last digits",
        "account last digits",
        "text extraction",
        "split text",
        "excel text function",
        "data cleaning",
        "data preparation",
        "automation",
        "advanced excel",
        "mis reporting",
        "string extraction",
        "text analysis",
        "text formula",
        "report preparation"
    ],

    aliases:[
        "ending text formula",
        "last letters formula",
        "suffix formula"
    ],

    priorityRules:[
        "If query contains last characters => RIGHT"
    ],

    questions:[
        "How to extract last characters?",
        "How to get ending text?",
        "How to use RIGHT formula?"
    ],

    dataset:"customer_data",

    category:"Text",

    formula:'=RIGHT(A2,4)',

    syntax:"RIGHT(text,[num_chars])",

    explanation:"Returns characters from the end of a text string.",

    breakdown:[
        "A2 = Text Value",
        "4 = Characters Required",
        "Returns Ending Characters"
    ],

    output:"4589",

    business:"Used for account numbers, IDs, and suffix extraction.",

    tips:[
        "Useful for extracting last digits.",
        "Combine with VALUE for numbers."
    ],

    errors:[
        "#VALUE!"
    ],

    alternative:'=TEXTAFTER(A2,"-")'
},


// ==========================================================
// MID
// ==========================================================
{
    title:"MID",

    keywords:[
        "mid",
        "middle characters",
        "extract middle text",
        "text from position",
        "extract characters",
        "employee code parsing",
        "id extraction",
        "substring",
        "text splitting",
        "fixed position extraction",
        "excel text function",
        "data cleaning",
        "data preparation",
        "automation",
        "advanced excel",
        "mis reporting",
        "string formula",
        "text analysis",
        "report preparation",
        "text extraction"
    ],

    aliases:[
        "substring formula",
        "middle text formula",
        "position extraction"
    ],

    priorityRules:[
        "If query contains middle characters => MID"
    ],

    questions:[
        "How to extract middle text?",
        "How to get characters from position?",
        "How to use MID formula?"
    ],

    dataset:"employee_id",

    category:"Text",

    formula:'=MID(A2,4,5)',

    syntax:"MID(text,start_num,num_chars)",

    explanation:"Returns characters from the middle of a text string based on position.",

    breakdown:[
        "A2 = Text Value",
        "4 = Starting Position",
        "5 = Number of Characters"
    ],

    output:"12345",

    business:"Used for extracting IDs, codes, and structured text data.",

    tips:[
        "Position starts from 1.",
        "Useful for fixed-format data."
    ],

    errors:[
        "#VALUE!"
    ],

    alternative:'=TEXTBEFORE(TEXTAFTER(A2,"-"),"-")'
},


// ==========================================================
// LEN
// ==========================================================
{
    title:"LEN",

    keywords:[
        "len",
        "count characters",
        "text length",
        "number of characters",
        "character count",
        "validate length",
        "check id length",
        "password length",
        "data validation",
        "text analysis",
        "excel text function",
        "data cleaning",
        "automation",
        "advanced excel",
        "mis reporting",
        "quality check",
        "string length",
        "text formula",
        "report preparation",
        "data validation"
    ],

    aliases:[
        "length formula",
        "character counter",
        "text count formula"
    ],

    priorityRules:[
        "If query contains count characters => LEN"
    ],

    questions:[
        "How to count characters?",
        "How to check text length?",
        "How to use LEN formula?"
    ],

    dataset:"customer_master",

    category:"Text",

    formula:"=LEN(A2)",

    syntax:"LEN(text)",

    explanation:"Returns the number of characters in a text string.",

    breakdown:[
        "A2 = Text Value",
        "Counts Letters",
        "Counts Spaces Also"
    ],

    output:"12",

    business:"Used for data validation and quality checks.",

    tips:[
        "Spaces are counted as characters.",
        "Use TRIM before LEN for clean count."
    ],

    errors:[
        "#VALUE!"
    ],

    alternative:"=LEN(TRIM(A2))"
},


// ==========================================================
// TRIM
// ==========================================================
{
    title:"TRIM",

    keywords:[
        "trim",
        "remove spaces",
        "extra spaces",
        "clean text",
        "remove unwanted spaces",
        "data cleaning",
        "fix imported data",
        "normalize text",
        "clean employee names",
        "remove leading spaces",
        "remove trailing spaces",
        "excel text function",
        "power query cleaning",
        "automation",
        "advanced excel",
        "mis reporting",
        "data preparation",
        "quality check",
        "text cleanup",
        "string cleaning"
    ],

    aliases:[
        "space remover",
        "text cleaner",
        "clean spaces formula"
    ],

    priorityRules:[
        "If query contains remove spaces => TRIM"
    ],

    questions:[
        "How to remove extra spaces?",
        "How to clean text?",
        "How to use TRIM formula?"
    ],

    dataset:"employee_master",

    category:"Text",

    formula:"=TRIM(A2)",

    syntax:"TRIM(text)",

    explanation:"Removes extra spaces from text except single spaces between words.",

    breakdown:[
        "A2 = Text Value",
        "Removes Leading Spaces",
        "Removes Extra Spaces"
    ],

    output:"Sunil Kumar",

    business:"Used for cleaning imported Excel data and master files.",

    tips:[
        "Use before lookup formulas.",
        "Improves matching accuracy."
    ],

    errors:[
        "#VALUE!"
    ],

    alternative:"=CLEAN(TRIM(A2))"
},

// ==========================================================
// CONCAT
// ==========================================================
{
    title:"CONCAT",

    keywords:[
        "concat",
        "combine text",
        "join text",
        "merge text",
        "combine columns",
        "combine cells",
        "create full name",
        "create employee id",
        "text combination",
        "merge data",
        "excel text function",
        "data preparation",
        "data cleaning",
        "automation",
        "advanced excel",
        "report preparation",
        "dashboard formula",
        "string combination",
        "text formula",
        "master data"
    ],

    aliases:[
        "text combine formula",
        "join cells formula",
        "merge columns"
    ],

    priorityRules:[
        "If query contains separator => TEXTJOIN",
        "If query contains old concatenate => CONCATENATE"
    ],

    questions:[
        "How to combine text?",
        "How to merge two columns?",
        "How to use CONCAT?"
    ],

    dataset:"employees",

    category:"Text",

    formula:"=CONCAT(A2,B2)",

    syntax:"CONCAT(text1,[text2],...)",

    explanation:"Combines text from multiple cells into one text string.",

    breakdown:[
        "A2 = First Name",
        "B2 = Last Name",
        "Returns Combined Text"
    ],

    output:"SunilKumar",

    business:"Used for creating names, IDs, and combined reporting fields.",

    tips:[
        "Use TEXTJOIN when separators are required.",
        "Replacement for CONCATENATE."
    ],

    errors:[
        "#VALUE!"
    ],

    alternative:'=TEXTJOIN(" ",TRUE,A2,B2)'
},


// ==========================================================
// CONCATENATE
// ==========================================================
{
    title:"CONCATENATE",

    keywords:[
        "concatenate",
        "old concatenate formula",
        "combine cells",
        "merge text",
        "join strings",
        "combine values",
        "create full name",
        "text joining",
        "legacy excel formula",
        "combine columns",
        "excel text function",
        "data preparation",
        "automation",
        "advanced excel",
        "report preparation",
        "string formula",
        "text manipulation",
        "data cleaning",
        "business reporting"
    ],

    aliases:[
        "concat old formula",
        "text merge formula",
        "join text old"
    ],

    priorityRules:[
        "If query contains old formula => CONCATENATE",
        "If query contains modern formula => CONCAT"
    ],

    questions:[
        "How to join text using CONCATENATE?",
        "How to combine values?",
        "How to use old concatenate formula?"
    ],

    dataset:"employees",

    category:"Text",

    formula:'=CONCATENATE(A2," ",B2)',

    syntax:"CONCATENATE(text1,text2,...)",

    explanation:"Joins multiple text strings into one text value.",

    breakdown:[
        "A2 = First Name",
        "\" \" = Space",
        "B2 = Last Name"
    ],

    output:"Sunil Kumar",

    business:"Used for creating labels, names, and combined fields.",

    tips:[
        "CONCAT is the modern replacement.",
        "TEXTJOIN is better for delimiters."
    ],

    errors:[
        "#VALUE!"
    ],

    alternative:'=CONCAT(A2," ",B2)'
},


// ==========================================================
// TEXTJOIN
// ==========================================================
{
    title:"TEXTJOIN",

    keywords:[
        "textjoin",
        "combine with separator",
        "merge with comma",
        "join text with delimiter",
        "combine multiple cells",
        "ignore blanks",
        "create sentence",
        "combine list",
        "merge records",
        "dynamic text",
        "excel text function",
        "automation",
        "advanced excel",
        "report preparation",
        "dashboard formula",
        "data preparation",
        "string manipulation",
        "business reporting",
        "text formula"
    ],

    aliases:[
        "delimiter combine formula",
        "comma join formula",
        "multiple text merge"
    ],

    priorityRules:[
        "If query contains separator => TEXTJOIN"
    ],

    questions:[
        "How to combine text with comma?",
        "How to merge multiple cells?",
        "How to ignore blanks while joining?"
    ],

    dataset:"employee_skills",

    category:"Text",

    formula:'=TEXTJOIN(", ",TRUE,A2:D2)',

    syntax:"TEXTJOIN(delimiter,ignore_empty,text1,...)",

    explanation:"Combines multiple text values using a specified separator.",

    breakdown:[
        "\", \" = Separator",
        "TRUE = Ignore Blank Cells",
        "A2:D2 = Text Range"
    ],

    output:"Excel,SQL,Power BI",

    business:"Used for skill lists, reports, and dynamic text creation.",

    tips:[
        "Best option for multiple cells.",
        "Avoids extra separators."
    ],

    errors:[
        "#VALUE!"
    ],

    alternative:'=CONCAT(A2:D2)'
},


// ==========================================================
// FIND
// ==========================================================
{
    title:"FIND",

    keywords:[
        "find",
        "find text position",
        "locate character",
        "text position",
        "case sensitive search",
        "character location",
        "extract based on position",
        "find symbol",
        "search character",
        "text analysis",
        "excel text function",
        "data cleaning",
        "automation",
        "advanced excel",
        "string analysis",
        "text extraction",
        "data preparation",
        "formula search"
    ],

    aliases:[
        "case sensitive search",
        "character finder",
        "position finder"
    ],

    priorityRules:[
        "If query contains case sensitive => FIND",
        "If query contains ignore case => SEARCH"
    ],

    questions:[
        "How to find text position?",
        "How to locate a character?",
        "How to use FIND?"
    ],

    dataset:"customer_data",

    category:"Text",

    formula:'=FIND("-",A2)',

    syntax:"FIND(find_text,within_text,[start_num])",

    explanation:"Returns the position of one text string inside another and is case-sensitive.",

    breakdown:[
        "\"-\" = Character To Find",
        "A2 = Text Cell",
        "Returns Position Number"
    ],

    output:"5",

    business:"Used for text extraction and structured data parsing.",

    tips:[
        "Case-sensitive function.",
        "Returns error if text is not found."
    ],

    errors:[
        "#VALUE!"
    ],

    alternative:'=SEARCH("-",A2)'
},

// ==========================================================
// ISBLANK
// ==========================================================
{
    title:"ISBLANK",

    keywords:[
        "isblank",
        "blank check",
        "check empty cell",
        "empty cell validation",
        "missing value check",
        "blank detection",
        "data completeness",
        "empty record check",
        "null value check",
        "data quality check",
        "excel information function",
        "data validation",
        "advanced excel",
        "mis reporting",
        "report automation",
        "quality control",
        "formula checking",
        "audit check",
        "data cleaning",
        "information formula"
    ],

    aliases:[
        "empty checker",
        "blank validator",
        "missing data formula"
    ],

    priorityRules:[
        "If query contains blank or empty => ISBLANK",
        "If query contains error => IFERROR"
    ],

    questions:[
        "How to check blank cells?",
        "How to identify missing data?",
        "How to use ISBLANK?"
    ],

    dataset:"employee_master",

    category:"Information",

    formula:"=ISBLANK(A2)",

    syntax:"ISBLANK(value)",

    explanation:"Checks whether a cell is empty and returns TRUE or FALSE.",

    breakdown:[
        "A2 = Cell To Check",
        "Empty Cell Returns TRUE",
        "Value Cell Returns FALSE"
    ],

    output:"FALSE",

    business:"Used for data quality checks and incomplete record monitoring.",

    tips:[
        "Use before reporting.",
        "Blank and zero are different."
    ],

    errors:[
        "#VALUE!"
    ],

    alternative:'=IF(A2="","Blank","Filled")'
},


// ==========================================================
// ISNUMBER
// ==========================================================
{
    title:"ISNUMBER",

    keywords:[
        "isnumber",
        "number check",
        "check numeric value",
        "identify numbers",
        "numeric validation",
        "number detection",
        "data type check",
        "formula validation",
        "numeric test",
        "excel information function",
        "data cleaning",
        "advanced excel",
        "mis reporting",
        "report automation",
        "quality check",
        "data analysis",
        "information formula",
        "validation formula",
        "number verification"
    ],

    aliases:[
        "numeric checker",
        "number validator",
        "numeric test formula"
    ],

    priorityRules:[
        "If query contains number => ISNUMBER",
        "If query contains text => ISTEXT"
    ],

    questions:[
        "How to check if value is number?",
        "How to validate numeric data?",
        "How to use ISNUMBER?"
    ],

    dataset:"sales_data",

    category:"Information",

    formula:"=ISNUMBER(A2)",

    syntax:"ISNUMBER(value)",

    explanation:"Checks whether a value is a number.",

    breakdown:[
        "A2 = Value To Test",
        "Number Returns TRUE",
        "Text Returns FALSE"
    ],

    output:"TRUE",

    business:"Used for validation and preventing calculation errors.",

    tips:[
        "Useful with SEARCH formulas.",
        "Combine with IF for validation."
    ],

    errors:[
        "#VALUE!"
    ],

    alternative:'=IF(ISNUMBER(A2),"Valid","Invalid")'
},


// ==========================================================
// ISTEXT
// ==========================================================
{
    title:"ISTEXT",

    keywords:[
        "istext",
        "text check",
        "check text value",
        "identify text",
        "text validation",
        "string check",
        "character validation",
        "data type checking",
        "text detection",
        "excel information function",
        "data cleaning",
        "advanced excel",
        "mis reporting",
        "report automation",
        "quality check",
        "data analysis",
        "information formula",
        "validation formula",
        "text verification"
    ],

    aliases:[
        "text checker",
        "string validator",
        "text test formula"
    ],

    priorityRules:[
        "If query contains text => ISTEXT",
        "If query contains number => ISNUMBER"
    ],

    questions:[
        "How to check text value?",
        "How to identify text cells?",
        "How to use ISTEXT?"
    ],

    dataset:"customer_master",

    category:"Information",

    formula:"=ISTEXT(A2)",

    syntax:"ISTEXT(value)",

    explanation:"Checks whether a value contains text.",

    breakdown:[
        "A2 = Value To Test",
        "Text Returns TRUE",
        "Number Returns FALSE"
    ],

    output:"TRUE",

    business:"Used for text validation and data cleaning.",

    tips:[
        "Useful before text processing.",
        "Combine with IF."
    ],

    errors:[
        "#VALUE!"
    ],

    alternative:'=IF(ISTEXT(A2),"Text","Not Text")'
},


// ==========================================================
// IFERROR
// ==========================================================
{
    title:"IFERROR",

    keywords:[
        "iferror",
        "handle error",
        "remove errors",
        "replace error",
        "avoid formula error",
        "error handling",
        "clean report",
        "hide errors",
        "formula protection",
        "replace #n/a",
        "replace div zero",
        "excel information function",
        "advanced excel",
        "mis reporting",
        "dashboard formula",
        "report automation",
        "data quality",
        "error management",
        "business reporting",
        "formula improvement"
    ],

    aliases:[
        "error handler",
        "error replacement formula",
        "safe formula"
    ],

    priorityRules:[
        "If query contains any error => IFERROR",
        "If query contains lookup missing only => IFNA"
    ],

    questions:[
        "How to handle formula errors?",
        "How to replace errors?",
        "How to use IFERROR?"
    ],

    dataset:"lookup_report",

    category:"Information",

    formula:'=IFERROR(VLOOKUP(A2,D:E,2,FALSE),"Not Found")',

    syntax:"IFERROR(value,value_if_error)",

    explanation:"Returns a custom result when a formula generates an error.",

    breakdown:[
        "Formula = Value To Check",
        "\"Not Found\" = Error Result",
        "Avoids Visible Errors"
    ],

    output:"Not Found",

    business:"Used for clean dashboards and professional reports.",

    tips:[
        "Improves report presentation.",
        "Do not hide errors during testing."
    ],

    errors:[
        "#VALUE!",
        "#REF!",
        "#DIV/0!"
    ],

    alternative:'=IFNA(VLOOKUP(A2,D:E,2,FALSE),"Not Found")'
},


// ==========================================================
// IFNA
// ==========================================================
{
    title:"IFNA",

    keywords:[
        "ifna",
        "replace n/a",
        "handle #n/a",
        "lookup missing value",
        "missing lookup result",
        "vlookup error handling",
        "xlookup error",
        "not available handling",
        "lookup protection",
        "missing record",
        "excel information function",
        "advanced excel",
        "mis reporting",
        "dashboard formula",
        "data validation",
        "report automation",
        "lookup formula",
        "error handling",
        "business reporting",
        "data quality"
    ],

    aliases:[
        "na error handler",
        "lookup error formula",
        "missing value handler"
    ],

    priorityRules:[
        "If query contains #N/A => IFNA",
        "If query contains all errors => IFERROR"
    ],

    questions:[
        "How to handle #N/A?",
        "How to replace lookup error?",
        "How to use IFNA?"
    ],

    dataset:"employee_lookup",

    category:"Information",

    formula:'=IFNA(XLOOKUP(A2,A:A,D:D),"Not Found")',

    syntax:"IFNA(value,value_if_na)",

    explanation:"Handles only #N/A errors and returns a custom value.",

    breakdown:[
        "XLOOKUP = Formula",
        "\"Not Found\" = Missing Result",
        "Handles #N/A Only"
    ],

    output:"Not Found",

    business:"Used for lookup reports where missing records are expected.",

    tips:[
        "Use IFNA for lookup formulas.",
        "Use IFERROR for broader error handling."
    ],

    errors:[
        "#N/A"
    ],

    alternative:'=IFERROR(XLOOKUP(A2,A:A,D:D),"Not Found")'
},

// ==========================================================
// FILTER
// ==========================================================
{
    title:"FILTER",

    keywords:[
        "filter",
        "filter data",
        "extract records",
        "dynamic filter",
        "return matching rows",
        "conditional extraction",
        "dynamic report",
        "live report",
        "filter table",
        "extract employee records",
        "excel dynamic array",
        "modern excel",
        "advanced excel",
        "dashboard formula",
        "data analysis",
        "mis reporting",
        "report automation",
        "business reporting",
        "spill formula",
        "excel 365"
    ],

    aliases:[
        "dynamic filter formula",
        "auto filter formula",
        "record extraction formula"
    ],

    priorityRules:[
        "If query contains remove duplicate => UNIQUE",
        "If query contains filter records => FILTER"
    ],

    questions:[
        "How to filter data dynamically?",
        "How to extract matching records?",
        "How to use FILTER formula?"
    ],

    dataset:"employees",

    category:"Dynamic Array",

    formula:'=FILTER(A2:D100,D2:D100="Active")',

    syntax:"FILTER(array,include,[if_empty])",

    explanation:"Returns only rows that meet specified conditions.",

    breakdown:[
        "A2:D100 = Data Range",
        "D2:D100 = Condition Column",
        "Active = Filter Criteria"
    ],

    output:"Active Employees List",

    business:"Used for dynamic MIS reports and live dashboards.",

    tips:[
        "Available in Excel 365.",
        "Automatically spills results."
    ],

    errors:[
        "#CALC!",
        "#VALUE!"
    ],

    alternative:'=IFERROR(FILTER(A:D,D:D="Active"),"No Data")'
},


// ==========================================================
// SORT
// ==========================================================
{
    title:"SORT",

    keywords:[
        "sort",
        "sort data",
        "ascending order",
        "descending order",
        "arrange data",
        "dynamic sorting",
        "auto sort",
        "sort table",
        "rank data",
        "organize records",
        "excel dynamic array",
        "modern excel",
        "advanced excel",
        "dashboard formula",
        "data analysis",
        "report automation",
        "business reporting",
        "excel 365",
        "spill formula"
    ],

    aliases:[
        "automatic sorting",
        "dynamic sort formula",
        "order data formula"
    ],

    priorityRules:[
        "If query contains based on another column => SORTBY"
    ],

    questions:[
        "How to sort data automatically?",
        "How to arrange records?",
        "How to use SORT formula?"
    ],

    dataset:"sales",

    category:"Dynamic Array",

    formula:"=SORT(A2:D100,4,-1)",

    syntax:"SORT(array,[sort_index],[sort_order])",

    explanation:"Sorts a range dynamically without changing original data.",

    breakdown:[
        "A2:D100 = Data Range",
        "4 = Sort Column",
        "-1 = Descending Order"
    ],

    output:"Sorted Sales Report",

    business:"Used for dynamic ranking and automated reports.",

    tips:[
        "Original data remains unchanged.",
        "Works with spilled arrays."
    ],

    errors:[
        "#VALUE!",
        "#SPILL!"
    ],

    alternative:"=SORTBY(A2:D100,D2:D100,-1)"
},


// ==========================================================
// SORTBY
// ==========================================================
{
    title:"SORTBY",

    keywords:[
        "sortby",
        "sort based on column",
        "custom sorting",
        "dynamic sorting",
        "sort using another range",
        "multiple sorting",
        "advanced sort",
        "conditional sorting",
        "rank report",
        "priority sorting",
        "excel dynamic array",
        "modern excel",
        "advanced excel",
        "dashboard formula",
        "data analysis",
        "report automation",
        "business reporting",
        "excel 365",
        "spill formula"
    ],

    aliases:[
        "custom sort formula",
        "multi column sorting",
        "advanced sorting"
    ],

    priorityRules:[
        "If query contains another column => SORTBY",
        "If simple sorting => SORT"
    ],

    questions:[
        "How to sort using another column?",
        "How to create custom sorting?",
        "How to use SORTBY?"
    ],

    dataset:"employee_performance",

    category:"Dynamic Array",

    formula:"=SORTBY(A2:D100,D2:D100,-1)",

    syntax:"SORTBY(array,by_array,[sort_order])",

    explanation:"Sorts one range based on values from another range.",

    breakdown:[
        "A2:D100 = Data",
        "D2:D100 = Sorting Column",
        "-1 = Descending"
    ],

    output:"Performance Ranking",

    business:"Used for ranking reports and dashboards.",

    tips:[
        "Supports multiple sorting levels.",
        "More flexible than SORT."
    ],

    errors:[
        "#VALUE!",
        "#SPILL!"
    ],

    alternative:"=SORT(A2:D100,4,-1)"
},


// ==========================================================
// UNIQUE
// ==========================================================
{
    title:"UNIQUE",

    keywords:[
        "unique",
        "remove duplicates",
        "distinct values",
        "duplicate removal",
        "get unique list",
        "deduplicate data",
        "master list creation",
        "customer list",
        "employee list",
        "clean database",
        "excel dynamic array",
        "modern excel",
        "advanced excel",
        "data cleaning",
        "data preparation",
        "mis reporting",
        "dashboard formula",
        "excel 365",
        "spill formula"
    ],

    aliases:[
        "duplicate remover",
        "distinct formula",
        "unique list formula"
    ],

    priorityRules:[
        "If query contains duplicate => UNIQUE"
    ],

    questions:[
        "How to remove duplicates?",
        "How to create unique list?",
        "How to use UNIQUE formula?"
    ],

    dataset:"customer_master",

    category:"Dynamic Array",

    formula:"=UNIQUE(A2:A100)",

    syntax:"UNIQUE(array,[by_col],[exactly_once])",

    explanation:"Returns a list of unique values from a range.",

    breakdown:[
        "A2:A100 = Source Data",
        "Removes Duplicate Values",
        "Returns Unique Records"
    ],

    output:"Unique Customer Names",

    business:"Used for master data creation and cleanup.",

    tips:[
        "Better than Remove Duplicates tool.",
        "Updates automatically."
    ],

    errors:[
        "#CALC!",
        "#SPILL!"
    ],

    alternative:"=SORT(UNIQUE(A2:A100))"
},


// ==========================================================
// SEQUENCE
// ==========================================================
{
    title:"SEQUENCE",

    keywords:[
        "sequence",
        "generate numbers",
        "auto numbering",
        "create serial number",
        "dynamic numbering",
        "number series",
        "automatic sequence",
        "generate dates",
        "create list",
        "array generation",
        "excel dynamic array",
        "modern excel",
        "advanced excel",
        "dashboard formula",
        "automation",
        "report preparation",
        "excel 365",
        "spill formula",
        "dynamic formula",
        "number generation"
    ],

    aliases:[
        "serial number formula",
        "auto sequence formula",
        "number generator"
    ],

    priorityRules:[
        "If query contains generate numbers => SEQUENCE"
    ],

    questions:[
        "How to generate serial numbers?",
        "How to create number series?",
        "How to use SEQUENCE formula?"
    ],

    dataset:"number_list",

    category:"Dynamic Array",

    formula:"=SEQUENCE(10)",

    syntax:"SEQUENCE(rows,[columns],[start],[step])",

    explanation:"Generates a sequence of numbers automatically.",

    breakdown:[
        "10 = Number of Rows",
        "Creates Automatic Series",
        "Spills Results"
    ],

    output:"1,2,3,4,5...",

    business:"Used for automation, numbering, and dynamic reports.",

    tips:[
        "Can generate dates also.",
        "Works only in Excel 365."
    ],

    errors:[
        "#SPILL!",
        "#VALUE!"
    ],

    alternative:"=ROW(A1:A10)"
},

// ==========================================================
// SWITCH
// ==========================================================
{
    title:"SWITCH",

    keywords:[
        "switch",
        "replace nested if",
        "multiple choice formula",
        "condition based result",
        "value mapping",
        "category formula",
        "grade formula",
        "status mapping",
        "scenario formula",
        "decision formula",
        "excel logical function",
        "advanced excel",
        "automation",
        "mis reporting",
        "dashboard formula",
        "business rules",
        "classification formula",
        "formula optimization",
        "logical formula",
        "excel 365"
    ],

    aliases:[
        "nested if alternative",
        "case formula",
        "multiple option formula"
    ],

    priorityRules:[
        "If query contains multiple conditions => IFS",
        "If query contains matching values => SWITCH"
    ],

    questions:[
        "How to replace nested IF?",
        "How to map multiple values?",
        "How to use SWITCH?"
    ],

    dataset:"employee_grade",

    category:"Logical",

    formula:'=SWITCH(A2,"A","Excellent","B","Good","C","Average")',

    syntax:"SWITCH(expression,value1,result1,...,default)",

    explanation:"Returns a result based on matching values.",

    breakdown:[
        "A2 = Value To Check",
        "A/B/C = Matching Options",
        "Returns Corresponding Result"
    ],

    output:"Excellent",

    business:"Used for employee grading, status mapping, and classification.",

    tips:[
        "Cleaner than multiple IF statements.",
        "Best when comparing one value."
    ],

    errors:[
        "#N/A",
        "#VALUE!"
    ],

    alternative:'=IF(A2="A","Excellent","Other")'
},


// ==========================================================
// IFS
// ==========================================================
{
    title:"IFS",

    keywords:[
        "ifs",
        "multiple conditions",
        "multiple criteria",
        "nested if replacement",
        "condition formula",
        "grading formula",
        "salary band",
        "performance rating",
        "decision formula",
        "logical test",
        "advanced if formula",
        "excel logical function",
        "advanced excel",
        "automation",
        "mis reporting",
        "dashboard formula",
        "business rules",
        "classification",
        "excel 365"
    ],

    aliases:[
        "multiple if formula",
        "condition based formula",
        "advanced if"
    ],

    priorityRules:[
        "If query contains multiple conditions => IFS",
        "If query contains single condition => IF"
    ],

    questions:[
        "How to handle multiple conditions?",
        "How to replace nested IF?",
        "How to use IFS?"
    ],

    dataset:"performance",

    category:"Logical",

    formula:'=IFS(B2>=90,"A",B2>=75,"B",B2>=60,"C",TRUE,"D")',

    syntax:"IFS(test1,result1,test2,result2,...)",

    explanation:"Checks multiple conditions and returns the first TRUE result.",

    breakdown:[
        "B2 = Score",
        "Conditions = Multiple Tests",
        "Returns First Matching Result"
    ],

    output:"A",

    business:"Used for grading, ratings, and business classifications.",

    tips:[
        "Conditions are evaluated from top to bottom.",
        "Add TRUE as default condition."
    ],

    errors:[
        "#N/A",
        "#VALUE!"
    ],

    alternative:'=IF(B2>=90,"A",IF(B2>=75,"B","C"))'
},


// ==========================================================
// AND
// ==========================================================
{
    title:"AND",

    keywords:[
        "and",
        "all conditions true",
        "multiple checks",
        "validate all criteria",
        "combined condition",
        "logical validation",
        "approval formula",
        "eligibility check",
        "compliance check",
        "rule validation",
        "excel logical function",
        "advanced excel",
        "automation",
        "mis reporting",
        "dashboard formula",
        "business rules",
        "data validation",
        "decision formula",
        "logical formula"
    ],

    aliases:[
        "all criteria formula",
        "multiple true condition",
        "combined test"
    ],

    priorityRules:[
        "If query contains all conditions => AND",
        "If query contains any condition => OR"
    ],

    questions:[
        "How to check multiple conditions?",
        "How to validate all rules?",
        "How to use AND?"
    ],

    dataset:"loan_application",

    category:"Logical",

    formula:"=AND(B2>=50000,C2=\"Yes\")",

    syntax:"AND(logical1,[logical2],...)",

    explanation:"Returns TRUE only when all conditions are TRUE.",

    breakdown:[
        "B2>=50000 = Condition 1",
        "C2=\"Yes\" = Condition 2",
        "Both Must Be TRUE"
    ],

    output:"TRUE",

    business:"Used for eligibility checks and approval workflows.",

    tips:[
        "Commonly used with IF.",
        "All conditions must pass."
    ],

    errors:[
        "#VALUE!"
    ],

    alternative:'=IF(AND(B2>=50000,C2="Yes"),"Approved","Rejected")'
},


// ==========================================================
// OR
// ==========================================================
{
    title:"OR",

    keywords:[
        "or",
        "any condition true",
        "multiple options",
        "alternative condition",
        "check one condition",
        "logical validation",
        "exception check",
        "approval rule",
        "risk check",
        "compliance condition",
        "excel logical function",
        "advanced excel",
        "automation",
        "mis reporting",
        "dashboard formula",
        "business rules",
        "decision formula",
        "logical formula",
        "data validation"
    ],

    aliases:[
        "any criteria formula",
        "one condition formula",
        "multiple option check"
    ],

    priorityRules:[
        "If query contains any condition => OR",
        "If query contains all conditions => AND"
    ],

    questions:[
        "How to check any condition?",
        "How to allow multiple options?",
        "How to use OR?"
    ],

    dataset:"employee_status",

    category:"Logical",

    formula:'=OR(A2="Active",A2="Pending")',

    syntax:"OR(logical1,[logical2],...)",

    explanation:"Returns TRUE when at least one condition is TRUE.",

    breakdown:[
        "Condition 1 = Active",
        "Condition 2 = Pending",
        "Any TRUE Returns TRUE"
    ],

    output:"TRUE",

    business:"Used for exception reports and multiple criteria checks.",

    tips:[
        "Useful with IF formulas.",
        "Only one condition needs to pass."
    ],

    errors:[
        "#VALUE!"
    ],

    alternative:'=IF(OR(A2="Active",A2="Pending"),"Open","Closed")'
},


// ==========================================================
// NOT
// ==========================================================
{
    title:"NOT",

    keywords:[
        "not",
        "reverse condition",
        "opposite result",
        "negative condition",
        "exclude condition",
        "not equal check",
        "invert logic",
        "exception formula",
        "validation formula",
        "logical reversal",
        "excel logical function",
        "advanced excel",
        "automation",
        "mis reporting",
        "dashboard formula",
        "business rules",
        "decision formula",
        "logical formula",
        "data validation"
    ],

    aliases:[
        "reverse logic formula",
        "opposite condition",
        "negation formula"
    ],

    priorityRules:[
        "If query contains opposite/reverse => NOT"
    ],

    questions:[
        "How to reverse a condition?",
        "How to check not equal?",
        "How to use NOT?"
    ],

    dataset:"employee_status",

    category:"Logical",

    formula:'=NOT(A2="Closed")',

    syntax:"NOT(logical)",

    explanation:"Reverses the logical value TRUE to FALSE and FALSE to TRUE.",

    breakdown:[
        "A2=\"Closed\" = Original Test",
        "NOT reverses Result",
        "Returns Opposite Value"
    ],

    output:"TRUE",

    business:"Used for exception checking and exclusion rules.",

    tips:[
        "Useful with AND and OR.",
        "Simplifies negative conditions."
    ],

    errors:[
        "#VALUE!"
    ],

    alternative:'=A2<>"Closed"'
},

// ==========================================================
// ROUND
// ==========================================================
{
    title:"ROUND",

    keywords:[
        "round",
        "round number",
        "decimal rounding",
        "normal rounding",
        "reduce decimals",
        "financial rounding",
        "amount rounding",
        "number formatting",
        "decimal adjustment",
        "currency calculation",
        "excel math function",
        "advanced excel",
        "mis reporting",
        "financial analysis",
        "dashboard formula",
        "calculation accuracy",
        "number cleanup",
        "formula optimization",
        "mathematical formula",
        "data preparation"
    ],

    aliases:[
        "decimal remover",
        "number round formula",
        "precision formula"
    ],

    priorityRules:[
        "If query contains always increase => ROUNDUP",
        "If query contains always decrease => ROUNDDOWN"
    ],

    questions:[
        "How to round numbers?",
        "How to reduce decimal places?",
        "How to use ROUND?"
    ],

    dataset:"financial_report",

    category:"Math",

    formula:"=ROUND(A2,2)",

    syntax:"ROUND(number,num_digits)",

    explanation:"Rounds a number to the specified number of digits.",

    breakdown:[
        "A2 = Number Value",
        "2 = Decimal Places",
        "Returns Rounded Value"
    ],

    output:"1250.75",

    business:"Used for financial reports, pricing, and MIS calculations.",

    tips:[
        "5 or more rounds upward.",
        "4 or less rounds downward."
    ],

    errors:[
        "#VALUE!"
    ],

    alternative:"=MROUND(A2,0.05)"
},


// ==========================================================
// ROUNDUP
// ==========================================================
{
    title:"ROUNDUP",

    keywords:[
        "roundup",
        "always increase",
        "increase decimals",
        "round higher",
        "ceiling rounding",
        "maximum rounding",
        "invoice rounding",
        "financial adjustment",
        "price calculation",
        "decimal increase",
        "excel math function",
        "advanced excel",
        "mis reporting",
        "financial analysis",
        "dashboard formula",
        "calculation formula",
        "number adjustment",
        "mathematical formula",
        "data preparation",
        "business calculation"
    ],

    aliases:[
        "higher rounding",
        "increase number formula",
        "always upward round"
    ],

    priorityRules:[
        "If query contains always increase => ROUNDUP"
    ],

    questions:[
        "How to always round up?",
        "How to increase decimal value?",
        "How to use ROUNDUP?"
    ],

    dataset:"invoice",

    category:"Math",

    formula:"=ROUNDUP(A2,0)",

    syntax:"ROUNDUP(number,num_digits)",

    explanation:"Rounds a number away from zero.",

    breakdown:[
        "A2 = Number",
        "0 = No Decimal",
        "Always Increases Value"
    ],

    output:"126",

    business:"Used for billing, quantity calculation, and pricing models.",

    tips:[
        "Always moves away from zero.",
        "Useful for minimum quantity rules."
    ],

    errors:[
        "#VALUE!"
    ],

    alternative:"=CEILING(A2,1)"
},


// ==========================================================
// ROUNDDOWN
// ==========================================================
{
    title:"ROUNDDOWN",

    keywords:[
        "rounddown",
        "always decrease",
        "reduce value",
        "round lower",
        "floor rounding",
        "remove decimals",
        "truncate number",
        "lower rounding",
        "financial calculation",
        "quantity calculation",
        "excel math function",
        "advanced excel",
        "mis reporting",
        "financial analysis",
        "dashboard formula",
        "number adjustment",
        "mathematical formula",
        "data preparation",
        "business calculation",
        "decimal handling"
    ],

    aliases:[
        "lower rounding",
        "decrease number formula",
        "always downward round"
    ],

    priorityRules:[
        "If query contains always decrease => ROUNDDOWN"
    ],

    questions:[
        "How to always round down?",
        "How to remove decimal values?",
        "How to use ROUNDDOWN?"
    ],

    dataset:"inventory",

    category:"Math",

    formula:"=ROUNDDOWN(A2,0)",

    syntax:"ROUNDDOWN(number,num_digits)",

    explanation:"Rounds a number toward zero.",

    breakdown:[
        "A2 = Number",
        "0 = No Decimal",
        "Always Reduces Value"
    ],

    output:"125",

    business:"Used for inventory, quantities, and conservative calculations.",

    tips:[
        "Does not increase value.",
        "Useful for whole unit calculations."
    ],

    errors:[
        "#VALUE!"
    ],

    alternative:"=FLOOR(A2,1)"
},


// ==========================================================
// ABS
// ==========================================================
{
    title:"ABS",

    keywords:[
        "abs",
        "absolute value",
        "remove negative sign",
        "ignore minus",
        "positive difference",
        "variance calculation",
        "error difference",
        "deviation analysis",
        "gap calculation",
        "distance between values",
        "excel math function",
        "advanced excel",
        "mis reporting",
        "financial analysis",
        "dashboard formula",
        "data comparison",
        "audit analysis",
        "mathematical formula",
        "business reporting",
        "data quality"
    ],

    aliases:[
        "positive value formula",
        "difference without sign",
        "minus remover"
    ],

    priorityRules:[
        "If query contains remove negative => ABS"
    ],

    questions:[
        "How to remove negative sign?",
        "How to find absolute difference?",
        "How to use ABS?"
    ],

    dataset:"variance_report",

    category:"Math",

    formula:"=ABS(A2-B2)",

    syntax:"ABS(number)",

    explanation:"Returns the absolute value of a number without its sign.",

    breakdown:[
        "A2-B2 = Difference",
        "ABS Removes Negative Sign",
        "Returns Positive Value"
    ],

    output:"500",

    business:"Used for variance reports and reconciliation analysis.",

    tips:[
        "Useful for comparing values.",
        "Common in finance reports."
    ],

    errors:[
        "#VALUE!"
    ],

    alternative:"=MAX(A2,B2)-MIN(A2,B2)"
},


// ==========================================================
// MOD
// ==========================================================
{
    title:"MOD",

    keywords:[
        "mod",
        "remainder",
        "division remainder",
        "left after division",
        "odd even check",
        "number grouping",
        "cycle calculation",
        "batch calculation",
        "split groups",
        "mathematical remainder",
        "excel math function",
        "advanced excel",
        "automation",
        "mis reporting",
        "dashboard formula",
        "number analysis",
        "data processing",
        "mathematical formula",
        "business calculation",
        "formula logic"
    ],

    aliases:[
        "remainder formula",
        "division leftover",
        "odd even formula"
    ],

    priorityRules:[
        "If query contains remainder => MOD"
    ],

    questions:[
        "How to find remainder?",
        "How to check odd or even?",
        "How to use MOD?"
    ],

    dataset:"employee_shift",

    category:"Math",

    formula:"=MOD(A2,2)",

    syntax:"MOD(number,divisor)",

    explanation:"Returns the remainder after division.",

    breakdown:[
        "A2 = Number",
        "2 = Divisor",
        "Returns Remaining Value"
    ],

    output:"1",

    business:"Used for batch processing, grouping, and odd-even checks.",

    tips:[
        "MOD(number,2) checks odd/even.",
        "Useful for cyclic calculations."
    ],

    errors:[
        "#DIV/0!",
        "#VALUE!"
    ],

    alternative:"=ISEVEN(A2)"
},

// ==========================================================
// SUMPRODUCT
// ==========================================================
{
    title:"SUMPRODUCT",

    keywords:[
        "sumproduct",
        "weighted calculation",
        "weighted average",
        "multiple condition sum",
        "conditional multiplication",
        "sum with criteria",
        "advanced sum formula",
        "calculate total amount",
        "quantity price calculation",
        "sales analysis",
        "excel math function",
        "advanced excel",
        "mis reporting",
        "financial analysis",
        "dashboard formula",
        "business reporting",
        "data analysis",
        "automation",
        "calculation formula",
        "array formula"
    ],

    aliases:[
        "weighted sum formula",
        "multiple criteria sum",
        "advanced multiplication formula"
    ],

    priorityRules:[
        "If query contains weighted average => SUMPRODUCT",
        "If query contains filtered total => SUBTOTAL"
    ],

    questions:[
        "How to calculate weighted average?",
        "How to multiply and sum values?",
        "How to use SUMPRODUCT?"
    ],

    dataset:"sales_data",

    category:"Math",

    formula:"=SUMPRODUCT(B2:B10,C2:C10)",

    syntax:"SUMPRODUCT(array1,array2,...)",

    explanation:"Multiplies corresponding values and returns the sum of products.",

    breakdown:[
        "B2:B10 = Quantity",
        "C2:C10 = Price",
        "Returns Total Amount"
    ],

    output:"250000",

    business:"Used for weighted analysis, sales calculation, and financial models.",

    tips:[
        "Works with multiple conditions.",
        "No helper column required."
    ],

    errors:[
        "#VALUE!",
        "#N/A"
    ],

    alternative:"=SUM(B2:B10*C2:C10)"
},


// ==========================================================
// SUBTOTAL
// ==========================================================
{
    title:"SUBTOTAL",

    keywords:[
        "subtotal",
        "filtered total",
        "sum filtered data",
        "visible rows total",
        "ignore hidden rows",
        "dynamic total",
        "filtered calculation",
        "excel filter total",
        "report total",
        "summary calculation",
        "excel math function",
        "advanced excel",
        "mis reporting",
        "dashboard formula",
        "data analysis",
        "business reporting",
        "automation",
        "report preparation",
        "table calculation",
        "formula"
    ],

    aliases:[
        "filter aware sum",
        "visible cells sum",
        "dynamic summary formula"
    ],

    priorityRules:[
        "If query contains filter => SUBTOTAL",
        "If query contains errors => AGGREGATE"
    ],

    questions:[
        "How to sum filtered data?",
        "How to ignore hidden rows?",
        "How to use SUBTOTAL?"
    ],

    dataset:"sales_report",

    category:"Math",

    formula:"=SUBTOTAL(9,B2:B100)",

    syntax:"SUBTOTAL(function_num,ref1,[ref2])",

    explanation:"Performs calculations on visible cells only.",

    breakdown:[
        "9 = SUM Function",
        "B2:B100 = Data Range",
        "Ignores Filtered Rows"
    ],

    output:"125000",

    business:"Used in filtered MIS reports and dashboards.",

    tips:[
        "9 means SUM.",
        "1-11 includes hidden rows handling differently."
    ],

    errors:[
        "#VALUE!"
    ],

    alternative:"=SUM(B2:B100)"
},


// ==========================================================
// AGGREGATE
// ==========================================================
{
    title:"AGGREGATE",

    keywords:[
        "aggregate",
        "ignore errors",
        "ignore hidden rows",
        "advanced subtotal",
        "filtered calculation",
        "error handling calculation",
        "dynamic summary",
        "advanced total",
        "calculate visible data",
        "excel aggregation",
        "excel math function",
        "advanced excel",
        "mis reporting",
        "dashboard formula",
        "data analysis",
        "business reporting",
        "automation",
        "formula optimization",
        "report calculation"
    ],

    aliases:[
        "advanced subtotal formula",
        "error ignoring formula",
        "smart calculation"
    ],

    priorityRules:[
        "If query contains ignore errors => AGGREGATE",
        "If query contains only filter => SUBTOTAL"
    ],

    questions:[
        "How to ignore errors while calculating?",
        "How to use AGGREGATE?",
        "How to calculate visible values?"
    ],

    dataset:"financial_report",

    category:"Math",

    formula:"=AGGREGATE(9,6,B2:B100)",

    syntax:"AGGREGATE(function_num,options,array)",

    explanation:"Performs calculations while ignoring errors, hidden rows, or nested subtotals.",

    breakdown:[
        "9 = SUM",
        "6 = Ignore Errors",
        "B2:B100 = Range"
    ],

    output:"500000",

    business:"Used for complex MIS and error-prone datasets.",

    tips:[
        "More powerful than SUBTOTAL.",
        "Useful for dashboards."
    ],

    errors:[
        "#VALUE!",
        "#NUM!"
    ],

    alternative:"=SUBTOTAL(9,B2:B100)"
},


// ==========================================================
// INT
// ==========================================================
{
    title:"INT",

    keywords:[
        "int",
        "remove decimal",
        "integer value",
        "whole number",
        "truncate number",
        "floor number",
        "decimal removal",
        "round down number",
        "convert decimal",
        "number cleaning",
        "excel math function",
        "advanced excel",
        "mis reporting",
        "financial analysis",
        "dashboard formula",
        "calculation formula",
        "data preparation",
        "number formatting",
        "mathematical formula"
    ],

    aliases:[
        "whole number formula",
        "decimal remover",
        "integer conversion"
    ],

    priorityRules:[
        "If query contains remove decimal => INT",
        "If query contains always decrease => ROUNDDOWN"
    ],

    questions:[
        "How to remove decimals?",
        "How to get whole number?",
        "How to use INT?"
    ],

    dataset:"calculation",

    category:"Math",

    formula:"=INT(A2)",

    syntax:"INT(number)",

    explanation:"Rounds a number down to the nearest integer.",

    breakdown:[
        "A2 = Decimal Number",
        "Removes Decimal Part",
        "Returns Whole Number"
    ],

    output:"125",

    business:"Used for age calculation, quantities, and whole number reporting.",

    tips:[
        "Always rounds toward negative infinity.",
        "Different from TRUNC for negatives."
    ],

    errors:[
        "#VALUE!"
    ],

    alternative:"=ROUNDDOWN(A2,0)"
},


// ==========================================================
// CEILING
// ==========================================================
{
    title:"CEILING",

    keywords:[
        "ceiling",
        "round to multiple",
        "nearest multiple",
        "increase to next value",
        "pricing rounding",
        "quantity rounding",
        "pack size calculation",
        "minimum value",
        "financial rounding",
        "business rounding",
        "excel math function",
        "advanced excel",
        "mis reporting",
        "financial analysis",
        "dashboard formula",
        "calculation formula",
        "number adjustment",
        "mathematical formula",
        "automation",
        "business calculation"
    ],

    aliases:[
        "multiple rounding formula",
        "next multiple formula",
        "upward multiple"
    ],

    priorityRules:[
        "If query contains multiple rounding => CEILING",
        "If query contains decimal increase => ROUNDUP"
    ],

    questions:[
        "How to round to nearest multiple?",
        "How to increase to next value?",
        "How to use CEILING?"
    ],

    dataset:"inventory",

    category:"Math",

    formula:"=CEILING(A2,10)",

    syntax:"CEILING(number,significance)",

    explanation:"Rounds a number up to the nearest specified multiple.",

    breakdown:[
        "A2 = Number",
        "10 = Multiple",
        "Returns Next Multiple"
    ],

    output:"130",

    business:"Used for packaging, pricing, and inventory calculations.",

    tips:[
        "Useful when values must meet minimum units.",
        "Works with multiples like 5,10,100."
    ],

    errors:[
        "#NUM!",
        "#VALUE!"
    ],

    alternative:"=ROUNDUP(A2/10,0)*10"
},

// ==========================================================
// COUNT
// ==========================================================
{
    title:"COUNT",

    keywords:[
        "count",
        "count numbers",
        "count numeric values",
        "number count",
        "calculate records",
        "count transactions",
        "count employees",
        "numeric data count",
        "excel counting function",
        "statistical function",
        "data analysis",
        "mis reporting",
        "dashboard formula",
        "kpi calculation",
        "report automation",
        "business reporting",
        "excel formula",
        "data validation",
        "statistics"
    ],

    aliases:[
        "numeric counter",
        "number counting formula",
        "record counter"
    ],

    priorityRules:[
        "If query contains filled cells => COUNTA",
        "If query contains blank cells => COUNTBLANK"
    ],

    questions:[
        "How to count numbers?",
        "How to count records?",
        "How to use COUNT?"
    ],

    dataset:"sales_data",

    category:"Statistical",

    formula:"=COUNT(A2:A100)",

    syntax:"COUNT(value1,[value2],...)",

    explanation:"Counts cells that contain numeric values.",

    breakdown:[
        "A2:A100 = Range",
        "Counts Only Numbers",
        "Ignores Text And Blank Cells"
    ],

    output:"85",

    business:"Used for transaction counts, employee IDs, and KPI reports.",

    tips:[
        "COUNT ignores text values.",
        "Use COUNTA for all non-empty cells."
    ],

    errors:[
        "#VALUE!"
    ],

    alternative:"=COUNTA(A2:A100)"
},


// ==========================================================
// COUNTA
// ==========================================================
{
    title:"COUNTA",

    keywords:[
        "counta",
        "count filled cells",
        "count non blank",
        "records count",
        "total entries",
        "filled data count",
        "employee count",
        "customer count",
        "database count",
        "excel counting function",
        "statistical function",
        "data analysis",
        "mis reporting",
        "dashboard formula",
        "kpi calculation",
        "report automation",
        "business reporting",
        "data validation",
        "statistics",
        "excel formula"
    ],

    aliases:[
        "non empty counter",
        "filled cell counter",
        "record counting formula"
    ],

    priorityRules:[
        "If query contains all filled records => COUNTA"
    ],

    questions:[
        "How to count filled cells?",
        "How to count records?",
        "How to use COUNTA?"
    ],

    dataset:"employee_master",

    category:"Statistical",

    formula:"=COUNTA(A2:A100)",

    syntax:"COUNTA(value1,[value2],...)",

    explanation:"Counts all cells that contain any type of data.",

    breakdown:[
        "A2:A100 = Range",
        "Counts Text And Numbers",
        "Ignores Only Blank Cells"
    ],

    output:"120",

    business:"Used for headcount, customer database, and record tracking.",

    tips:[
        "Counts formulas returning empty text as non-empty.",
        "Useful for data completeness."
    ],

    errors:[
        "#VALUE!"
    ],

    alternative:"=ROWS(A2:A100)-COUNTBLANK(A2:A100)"
},


// ==========================================================
// COUNTBLANK
// ==========================================================
{
    title:"COUNTBLANK",

    keywords:[
        "countblank",
        "count empty cells",
        "blank cell count",
        "missing data count",
        "empty records",
        "incomplete data",
        "data quality check",
        "missing information",
        "blank validation",
        "excel counting function",
        "statistical function",
        "data cleaning",
        "mis reporting",
        "dashboard formula",
        "quality control",
        "report automation",
        "business reporting",
        "statistics",
        "excel formula"
    ],

    aliases:[
        "empty cell counter",
        "missing data formula",
        "blank checker"
    ],

    priorityRules:[
        "If query contains missing data => COUNTBLANK"
    ],

    questions:[
        "How to count blank cells?",
        "How to find missing records?",
        "How to use COUNTBLANK?"
    ],

    dataset:"customer_master",

    category:"Statistical",

    formula:"=COUNTBLANK(A2:A100)",

    syntax:"COUNTBLANK(range)",

    explanation:"Counts the number of empty cells in a range.",

    breakdown:[
        "A2:A100 = Range",
        "Checks Empty Cells",
        "Returns Blank Count"
    ],

    output:"15",

    business:"Used for data quality monitoring and incomplete forms.",

    tips:[
        "Useful before analysis.",
        "Helps identify missing information."
    ],

    errors:[
        "#VALUE!"
    ],

    alternative:'=COUNTIF(A2:A100,"")'
},


// ==========================================================
// COUNTIF
// ==========================================================
{
    title:"COUNTIF",

    keywords:[
        "countif",
        "count with condition",
        "single condition count",
        "conditional counting",
        "count matching values",
        "count employees",
        "count status",
        "count category",
        "criteria based count",
        "excel counting function",
        "statistical function",
        "advanced excel",
        "mis reporting",
        "dashboard formula",
        "kpi calculation",
        "business reporting",
        "data analysis",
        "automation",
        "statistics",
        "excel formula"
    ],

    aliases:[
        "condition counter",
        "criteria count formula",
        "single criteria counting"
    ],

    priorityRules:[
        "If query contains multiple conditions => COUNTIFS"
    ],

    questions:[
        "How to count based on condition?",
        "How many employees are active?",
        "How to use COUNTIF?"
    ],

    dataset:"employee_status",

    category:"Statistical",

    formula:'=COUNTIF(B2:B100,"Active")',

    syntax:"COUNTIF(range,criteria)",

    explanation:"Counts cells that meet a specified condition.",

    breakdown:[
        "B2:B100 = Status Range",
        "Active = Condition",
        "Returns Matching Count"
    ],

    output:"75",

    business:"Used for employee status, attendance, and KPI reports.",

    tips:[
        "Supports wildcards.",
        "Use COUNTIFS for multiple criteria."
    ],

    errors:[
        "#VALUE!"
    ],

    alternative:'=SUMPRODUCT(--(B2:B100="Active"))'
},


// ==========================================================
// COUNTIFS
// ==========================================================
{
    title:"COUNTIFS",

    keywords:[
        "countifs",
        "multiple condition count",
        "count multiple criteria",
        "advanced count",
        "conditional counting",
        "employee analysis",
        "attendance count",
        "sales count",
        "multiple filters count",
        "criteria based analysis",
        "excel counting function",
        "statistical function",
        "advanced excel",
        "mis reporting",
        "dashboard formula",
        "kpi calculation",
        "business reporting",
        "data analysis",
        "automation",
        "statistics"
    ],

    aliases:[
        "multi criteria counter",
        "advanced count formula",
        "multiple criteria counting"
    ],

    priorityRules:[
        "If query contains more than one condition => COUNTIFS"
    ],

    questions:[
        "How to count multiple conditions?",
        "How many employees meet criteria?",
        "How to use COUNTIFS?"
    ],

    dataset:"employee_report",

    category:"Statistical",

    formula:'=COUNTIFS(B2:B100,"Active",C2:C100,">5")',

    syntax:"COUNTIFS(criteria_range1,criteria1,...)",

    explanation:"Counts records that satisfy multiple conditions.",

    breakdown:[
        "B2:B100 = Status Range",
        "Active = First Condition",
        "C2:C100 = Second Condition"
    ],

    output:"35",

    business:"Used for HR analytics, compliance, and MIS dashboards.",

    tips:[
        "All conditions must be TRUE.",
        "Supports multiple criteria."
    ],

    errors:[
        "#VALUE!"
    ],

    alternative:'=SUMPRODUCT((B2:B100="Active")*(C2:C100>5))'
},

// ==========================================================
// AVERAGE
// ==========================================================
{
    title:"AVERAGE",

    keywords:[
        "average",
        "calculate average",
        "mean value",
        "average calculation",
        "average salary",
        "average sales",
        "average score",
        "performance average",
        "kpi average",
        "excel statistical function",
        "data analysis",
        "mis reporting",
        "dashboard formula",
        "business reporting",
        "trend analysis",
        "statistics",
        "excel formula",
        "calculation",
        "analytics"
    ],

    aliases:[
        "mean formula",
        "normal average",
        "simple average"
    ],

    priorityRules:[
        "If query contains condition => AVERAGEIF",
        "If query contains multiple conditions => AVERAGEIFS"
    ],

    questions:[
        "How to calculate average?",
        "How to find average value?",
        "How to use AVERAGE?"
    ],

    dataset:"sales_data",

    category:"Statistical",

    formula:"=AVERAGE(B2:B100)",

    syntax:"AVERAGE(number1,[number2],...)",

    explanation:"Calculates the arithmetic average of numbers.",

    breakdown:[
        "B2:B100 = Data Range",
        "Adds All Values",
        "Divides By Count"
    ],

    output:"2500",

    business:"Used for average sales, salary analysis, and KPI reporting.",

    tips:[
        "Ignores text values.",
        "Blank cells are ignored."
    ],

    errors:[
        "#DIV/0!",
        "#VALUE!"
    ],

    alternative:"=SUM(B2:B100)/COUNT(B2:B100)"
},


// ==========================================================
// AVERAGEIF
// ==========================================================
{
    title:"AVERAGEIF",

    keywords:[
        "averageif",
        "average with condition",
        "conditional average",
        "average based on criteria",
        "department average",
        "employee average",
        "sales average",
        "category average",
        "single criteria average",
        "excel statistical function",
        "data analysis",
        "mis reporting",
        "dashboard formula",
        "business reporting",
        "kpi calculation",
        "statistics",
        "excel formula",
        "analytics",
        "report automation"
    ],

    aliases:[
        "criteria average",
        "condition average formula",
        "single condition average"
    ],

    priorityRules:[
        "If query contains multiple conditions => AVERAGEIFS"
    ],

    questions:[
        "How to calculate average with condition?",
        "How to find department average?",
        "How to use AVERAGEIF?"
    ],

    dataset:"employee_salary",

    category:"Statistical",

    formula:'=AVERAGEIF(B2:B100,"IT",C2:C100)',

    syntax:"AVERAGEIF(range,criteria,average_range)",

    explanation:"Calculates average for cells matching a condition.",

    breakdown:[
        "B2:B100 = Criteria Range",
        "IT = Condition",
        "C2:C100 = Average Range"
    ],

    output:"75000",

    business:"Used for department-wise salary and performance analysis.",

    tips:[
        "Use one condition only.",
        "For multiple conditions use AVERAGEIFS."
    ],

    errors:[
        "#DIV/0!",
        "#VALUE!"
    ],

    alternative:'=SUMIF(B:B,"IT",C:C)/COUNTIF(B:B,"IT")'
},


// ==========================================================
// AVERAGEIFS
// ==========================================================
{
    title:"AVERAGEIFS",

    keywords:[
        "averageifs",
        "multiple condition average",
        "advanced average",
        "average multiple criteria",
        "department location average",
        "performance average",
        "salary analysis",
        "conditional average",
        "criteria based average",
        "excel statistical function",
        "data analysis",
        "mis reporting",
        "dashboard formula",
        "business reporting",
        "kpi calculation",
        "statistics",
        "excel formula",
        "analytics",
        "report automation"
    ],

    aliases:[
        "multi criteria average",
        "advanced average formula",
        "multiple filter average"
    ],

    priorityRules:[
        "If query contains multiple criteria => AVERAGEIFS"
    ],

    questions:[
        "How to average multiple conditions?",
        "How to calculate filtered average?",
        "How to use AVERAGEIFS?"
    ],

    dataset:"employee_performance",

    category:"Statistical",

    formula:'=AVERAGEIFS(D2:D100,B2:B100,"IT",C2:C100,">80")',

    syntax:"AVERAGEIFS(average_range,criteria_range1,criteria1,...)",

    explanation:"Calculates average when multiple conditions are satisfied.",

    breakdown:[
        "D2:D100 = Average Range",
        "IT = Department Condition",
        ">80 = Performance Condition"
    ],

    output:"88",

    business:"Used for advanced HR analytics and KPI reports.",

    tips:[
        "All conditions must match.",
        "Works with multiple criteria."
    ],

    errors:[
        "#DIV/0!",
        "#VALUE!"
    ],

    alternative:'=SUMIFS(D:D,B:B,"IT",C:C,">80")/COUNTIFS(B:B,"IT",C:C,">80")'
},


// ==========================================================
// MEDIAN
// ==========================================================
{
    title:"MEDIAN",

    keywords:[
        "median",
        "middle value",
        "middle number",
        "central value",
        "salary median",
        "remove outlier impact",
        "statistical analysis",
        "data distribution",
        "average alternative",
        "excel statistical function",
        "data analysis",
        "mis reporting",
        "dashboard formula",
        "business reporting",
        "trend analysis",
        "statistics",
        "analytics",
        "excel formula"
    ],

    aliases:[
        "middle calculation",
        "central tendency formula",
        "data midpoint"
    ],

    priorityRules:[
        "If query contains middle value => MEDIAN"
    ],

    questions:[
        "How to find middle value?",
        "How to calculate median?",
        "How to use MEDIAN?"
    ],

    dataset:"salary_data",

    category:"Statistical",

    formula:"=MEDIAN(B2:B100)",

    syntax:"MEDIAN(number1,[number2],...)",

    explanation:"Returns the middle value in a dataset.",

    breakdown:[
        "B2:B100 = Data Range",
        "Sorts Values Internally",
        "Returns Middle Number"
    ],

    output:"65000",

    business:"Used in salary analysis where extreme values affect average.",

    tips:[
        "Less affected by outliers than AVERAGE.",
        "Useful for salary benchmarking."
    ],

    errors:[
        "#NUM!",
        "#VALUE!"
    ],

    alternative:"=PERCENTILE.INC(B2:B100,0.5)"
},


// ==========================================================
// MODE
// ==========================================================
{
    title:"MODE",

    keywords:[
        "mode",
        "most common value",
        "most repeated number",
        "frequent value",
        "highest frequency",
        "common data point",
        "statistical analysis",
        "trend identification",
        "pattern analysis",
        "excel statistical function",
        "data analysis",
        "mis reporting",
        "dashboard formula",
        "business reporting",
        "statistics",
        "analytics",
        "excel formula"
    ],

    aliases:[
        "frequency formula",
        "repeated value formula",
        "most occurring value"
    ],

    priorityRules:[
        "If query contains repeated value => MODE"
    ],

    questions:[
        "How to find most repeated value?",
        "How to calculate MODE?",
        "How to identify common value?"
    ],

    dataset:"transaction_data",

    category:"Statistical",

    formula:"=MODE.SNGL(B2:B100)",

    syntax:"MODE.SNGL(number1,[number2],...)",

    explanation:"Returns the value that appears most frequently in a dataset.",

    breakdown:[
        "B2:B100 = Data Range",
        "Counts Frequency",
        "Returns Highest Occurring Value"
    ],

    output:"1000",

    business:"Used for trend analysis and identifying common transaction values.",

    tips:[
        "MODE.SNGL returns one value.",
        "MODE.MULT returns multiple modes."
    ],

    errors:[
        "#N/A",
        "#VALUE!"
    ],

    alternative:"=MODE.MULT(B2:B100)"
},

// ==========================================================
// TODAY
// ==========================================================
{
    title:"TODAY",

    keywords:[
        "today",
        "today date",
        "current date",
        "system date",
        "date formula",
        "calculate current date",
        "aging calculation",
        "sla calculation",
        "due date",
        "deadline tracking",
        "excel date function",
        "time intelligence",
        "mis reporting",
        "dashboard formula",
        "report automation",
        "business reporting",
        "date analysis",
        "excel formula",
        "automation",
        "daily report"
    ],

    aliases:[
        "current day formula",
        "automatic date",
        "live date formula"
    ],

    priorityRules:[
        "If query contains time also => NOW"
    ],

    questions:[
        "How to get today's date?",
        "How to calculate aging?",
        "How to use TODAY?"
    ],

    dataset:"task_tracker",

    category:"Date & Time",

    formula:"=TODAY()",

    syntax:"TODAY()",

    explanation:"Returns the current date automatically.",

    breakdown:[
        "No Input Required",
        "Fetches System Date",
        "Updates Automatically"
    ],

    output:"28-Jul-2026",

    business:"Used for SLA tracking, aging reports, and daily MIS.",

    tips:[
        "Updates when workbook recalculates.",
        "No arguments required."
    ],

    errors:[
        "#VALUE!"
    ],

    alternative:"=NOW()-TODAY()"
},


// ==========================================================
// NOW
// ==========================================================
{
    title:"NOW",

    keywords:[
        "now",
        "current date time",
        "current timestamp",
        "live time",
        "date and time",
        "system time",
        "time tracking",
        "timestamp formula",
        "login time",
        "report timestamp",
        "excel date function",
        "time intelligence",
        "mis reporting",
        "dashboard formula",
        "automation",
        "business reporting",
        "date analysis",
        "excel formula",
        "tracking"
    ],

    aliases:[
        "timestamp formula",
        "live date time",
        "current datetime"
    ],

    priorityRules:[
        "If query contains only date => TODAY"
    ],

    questions:[
        "How to get current date and time?",
        "How to create timestamp?",
        "How to use NOW?"
    ],

    dataset:"activity_log",

    category:"Date & Time",

    formula:"=NOW()",

    syntax:"NOW()",

    explanation:"Returns the current date and time.",

    breakdown:[
        "No Input Required",
        "Returns Date",
        "Returns Current Time"
    ],

    output:"28-Jul-2026 21:00",

    business:"Used for audit logs, tracking, and automated reports.",

    tips:[
        "Updates automatically.",
        "Useful for timestamps."
    ],

    errors:[
        "#VALUE!"
    ],

    alternative:"=TODAY()"
},


// ==========================================================
// DATE
// ==========================================================
{
    title:"DATE",

    keywords:[
        "date",
        "create date",
        "build date",
        "combine year month day",
        "generate date",
        "date creation",
        "convert numbers to date",
        "custom date",
        "deadline formula",
        "due date calculation",
        "excel date function",
        "time intelligence",
        "mis reporting",
        "dashboard formula",
        "automation",
        "business reporting",
        "date analysis",
        "excel formula"
    ],

    aliases:[
        "date generator",
        "manual date formula",
        "date builder"
    ],

    priorityRules:[
        "If query contains extract year => YEAR",
        "If query contains extract month => MONTH"
    ],

    questions:[
        "How to create date from values?",
        "How to combine year month day?",
        "How to use DATE?"
    ],

    dataset:"employee_joining",

    category:"Date & Time",

    formula:"=DATE(2026,7,28)",

    syntax:"DATE(year,month,day)",

    explanation:"Creates a valid Excel date from year, month, and day values.",

    breakdown:[
        "2026 = Year",
        "7 = Month",
        "28 = Day"
    ],

    output:"28-Jul-2026",

    business:"Used for joining dates, deadlines, and reporting periods.",

    tips:[
        "Better than typing dates manually.",
        "Works with date calculations."
    ],

    errors:[
        "#VALUE!",
        "#NUM!"
    ],

    alternative:"=TODAY()"
},


// ==========================================================
// YEAR
// ==========================================================
{
    title:"YEAR",

    keywords:[
        "year",
        "extract year",
        "get year from date",
        "date year",
        "financial year",
        "year wise report",
        "annual analysis",
        "date breakdown",
        "year calculation",
        "excel date function",
        "time intelligence",
        "mis reporting",
        "dashboard formula",
        "business reporting",
        "date analysis",
        "excel formula",
        "automation"
    ],

    aliases:[
        "year extraction formula",
        "date year formula",
        "annual formula"
    ],

    priorityRules:[
        "If query contains month => MONTH"
    ],

    questions:[
        "How to extract year from date?",
        "How to create yearly report?",
        "How to use YEAR?"
    ],

    dataset:"transaction_data",

    category:"Date & Time",

    formula:"=YEAR(A2)",

    syntax:"YEAR(serial_number)",

    explanation:"Returns the year from a date.",

    breakdown:[
        "A2 = Date Value",
        "Extracts Year",
        "Returns Four Digit Year"
    ],

    output:"2026",

    business:"Used for yearly MIS, financial analysis, and trends.",

    tips:[
        "Works with valid Excel dates.",
        "Useful with Pivot reports."
    ],

    errors:[
        "#VALUE!"
    ],

    alternative:'=TEXT(A2,"YYYY")'
},


// ==========================================================
// MONTH
// ==========================================================
{
    title:"MONTH",

    keywords:[
        "month",
        "extract month",
        "get month from date",
        "month wise report",
        "monthly analysis",
        "date month",
        "monthly MIS",
        "period analysis",
        "month calculation",
        "excel date function",
        "time intelligence",
        "mis reporting",
        "dashboard formula",
        "business reporting",
        "date analysis",
        "excel formula",
        "automation"
    ],

    aliases:[
        "month extraction formula",
        "monthly formula",
        "date month formula"
    ],

    priorityRules:[
        "If query contains year => YEAR"
    ],

    questions:[
        "How to extract month?",
        "How to create monthly report?",
        "How to use MONTH?"
    ],

    dataset:"sales_report",

    category:"Date & Time",

    formula:"=MONTH(A2)",

    syntax:"MONTH(serial_number)",

    explanation:"Returns the month number from a date.",

    breakdown:[
        "A2 = Date Value",
        "Extracts Month",
        "Returns Number 1-12"
    ],

    output:"7",

    business:"Used for monthly MIS reports and trend analysis.",

    tips:[
        "Combine with TEXT for month names.",
        "Useful in Pivot analysis."
    ],

    errors:[
        "#VALUE!"
    ],

    alternative:'=TEXT(A2,"MMMM")'
},

// ==========================================================
// DAY
// ==========================================================
{
    title:"DAY",

    keywords:[
        "day",
        "extract day",
        "get day from date",
        "date day",
        "day number",
        "daily report",
        "date breakdown",
        "day calculation",
        "excel date function",
        "time intelligence",
        "mis reporting",
        "dashboard formula",
        "business reporting",
        "date analysis",
        "excel formula",
        "automation",
        "calendar analysis"
    ],

    aliases:[
        "day extraction formula",
        "date day formula",
        "daily formula"
    ],

    priorityRules:[
        "If query contains month => MONTH",
        "If query contains year => YEAR"
    ],

    questions:[
        "How to extract day from date?",
        "How to get day number?",
        "How to use DAY?"
    ],

    dataset:"transaction_data",

    category:"Date & Time",

    formula:"=DAY(A2)",

    syntax:"DAY(serial_number)",

    explanation:"Returns the day number from a date.",

    breakdown:[
        "A2 = Date Value",
        "Extracts Day",
        "Returns Number 1-31"
    ],

    output:"28",

    business:"Used for daily analysis and date-based reporting.",

    tips:[
        "Works with valid Excel dates.",
        "Useful with MONTH and YEAR."
    ],

    errors:[
        "#VALUE!"
    ],

    alternative:'=TEXT(A2,"DD")'
},


// ==========================================================
// DATEDIF
// ==========================================================
{
    title:"DATEDIF",

    keywords:[
        "datedif",
        "age calculation",
        "calculate age",
        "date difference",
        "years between dates",
        "months between dates",
        "employee tenure",
        "service period",
        "duration calculation",
        "experience calculation",
        "excel date function",
        "time intelligence",
        "hr analytics",
        "mis reporting",
        "dashboard formula",
        "business reporting",
        "excel formula",
        "automation"
    ],

    aliases:[
        "age formula",
        "tenure formula",
        "date duration formula"
    ],

    priorityRules:[
        "If query contains age => DATEDIF",
        "If query contains working days => NETWORKDAYS"
    ],

    questions:[
        "How to calculate age?",
        "How to calculate experience?",
        "How to use DATEDIF?"
    ],

    dataset:"employee_master",

    category:"Date & Time",

    formula:'=DATEDIF(A2,TODAY(),"Y")',

    syntax:"DATEDIF(start_date,end_date,unit)",

    explanation:"Calculates the difference between two dates in years, months, or days.",

    breakdown:[
        "A2 = Joining/Birth Date",
        "TODAY() = Current Date",
        "Y = Complete Years"
    ],

    output:"8",

    business:"Used for employee age, tenure, and service period calculation.",

    tips:[
        "Hidden Excel function but widely used.",
        "Supports Y, M, and D units."
    ],

    errors:[
        "#NUM!",
        "#VALUE!"
    ],

    alternative:"=YEARFRAC(A2,TODAY())"
},


// ==========================================================
// DAYS
// ==========================================================
{
    title:"DAYS",

    keywords:[
        "days",
        "date difference",
        "number of days",
        "days between dates",
        "calculate duration",
        "deadline tracking",
        "aging days",
        "pending days",
        "project duration",
        "date calculation",
        "excel date function",
        "time intelligence",
        "mis reporting",
        "dashboard formula",
        "business reporting",
        "excel formula",
        "automation"
    ],

    aliases:[
        "date gap formula",
        "days calculation formula",
        "duration formula"
    ],

    priorityRules:[
        "If query contains working days => NETWORKDAYS"
    ],

    questions:[
        "How to calculate days between dates?",
        "How many days passed?",
        "How to use DAYS?"
    ],

    dataset:"ticket_tracker",

    category:"Date & Time",

    formula:"=DAYS(B2,A2)",

    syntax:"DAYS(end_date,start_date)",

    explanation:"Returns the number of days between two dates.",

    breakdown:[
        "B2 = End Date",
        "A2 = Start Date",
        "Returns Day Difference"
    ],

    output:"45",

    business:"Used for aging reports, SLA tracking, and project monitoring.",

    tips:[
        "Includes weekends.",
        "Use NETWORKDAYS for business days."
    ],

    errors:[
        "#VALUE!"
    ],

    alternative:"=B2-A2"
},


// ==========================================================
// NETWORKDAYS
// ==========================================================
{
    title:"NETWORKDAYS",

    keywords:[
        "networkdays",
        "working days",
        "business days",
        "exclude weekends",
        "exclude holidays",
        "sla days",
        "leave calculation",
        "office days",
        "project working days",
        "turnaround time",
        "excel date function",
        "time intelligence",
        "hr analytics",
        "mis reporting",
        "dashboard formula",
        "business reporting",
        "automation",
        "excel formula"
    ],

    aliases:[
        "working day calculator",
        "business day formula",
        "office days formula"
    ],

    priorityRules:[
        "If query contains working days => NETWORKDAYS",
        "If future date calculation => WORKDAY"
    ],

    questions:[
        "How to calculate working days?",
        "How to exclude weekends?",
        "How to calculate SLA days?"
    ],

    dataset:"leave_tracker",

    category:"Date & Time",

    formula:"=NETWORKDAYS(A2,B2)",

    syntax:"NETWORKDAYS(start_date,end_date,[holidays])",

    explanation:"Calculates working days excluding weekends and optional holidays.",

    breakdown:[
        "A2 = Start Date",
        "B2 = End Date",
        "Excludes Saturday/Sunday"
    ],

    output:"22",

    business:"Used for SLA, leave, productivity, and turnaround analysis.",

    tips:[
        "Add holiday range for accurate calculation.",
        "Useful for HR reports."
    ],

    errors:[
        "#VALUE!"
    ],

    alternative:"=DAYS(B2,A2)-WEEKENDS"
},


// ==========================================================
// WORKDAY
// ==========================================================
{
    title:"WORKDAY",

    keywords:[
        "workday",
        "future working date",
        "calculate due date",
        "business date",
        "exclude weekends",
        "project deadline",
        "sla due date",
        "delivery date",
        "next working day",
        "date planning",
        "excel date function",
        "time intelligence",
        "mis reporting",
        "dashboard formula",
        "business reporting",
        "automation",
        "excel formula"
    ],

    aliases:[
        "working date formula",
        "deadline formula",
        "business calendar formula"
    ],

    priorityRules:[
        "If query contains future date => WORKDAY",
        "If query contains working days between => NETWORKDAYS"
    ],

    questions:[
        "How to calculate future working date?",
        "How to find due date?",
        "How to use WORKDAY?"
    ],

    dataset:"project_tracker",

    category:"Date & Time",

    formula:"=WORKDAY(A2,10)",

    syntax:"WORKDAY(start_date,days,[holidays])",

    explanation:"Returns a future or past date excluding weekends and holidays.",

    breakdown:[
        "A2 = Start Date",
        "10 = Working Days",
        "Returns Due Date"
    ],

    output:"11-Aug-2026",

    business:"Used for project deadlines, SLA monitoring, and planning.",

    tips:[
        "Add holiday list for accuracy.",
        "Useful for automated reminders."
    ],

    errors:[
        "#VALUE!",
        "#NUM!"
    ],

    alternative:"=A2+10"
},

// ==========================================================
// LEFT
// ==========================================================
{
    title:"LEFT",

    keywords:[
        "left",
        "extract first characters",
        "get starting text",
        "first letters",
        "extract code",
        "employee code extraction",
        "prefix extraction",
        "text split",
        "data cleaning",
        "text processing",
        "excel text function",
        "advanced excel",
        "mis reporting",
        "dashboard formula",
        "data preparation",
        "automation",
        "business reporting",
        "text analysis",
        "excel formula"
    ],

    aliases:[
        "beginning text formula",
        "first character formula",
        "prefix formula"
    ],

    priorityRules:[
        "If query contains last characters => RIGHT",
        "If query contains middle characters => MID"
    ],

    questions:[
        "How to extract first characters?",
        "How to get starting text?",
        "How to use LEFT?"
    ],

    dataset:"employee_codes",

    category:"Text",

    formula:'=LEFT(A2,3)',

    syntax:"LEFT(text,[num_chars])",

    explanation:"Extracts characters from the beginning of a text string.",

    breakdown:[
        "A2 = Source Text",
        "3 = Number Of Characters",
        "Returns Left Side Text"
    ],

    output:"EMP",

    business:"Used for extracting employee codes, prefixes, and categories.",

    tips:[
        "Useful for fixed-length codes.",
        "Combine with LEN for dynamic extraction."
    ],

    errors:[
        "#VALUE!"
    ],

    alternative:"=TEXTBEFORE(A2,\"-\")"
},


// ==========================================================
// RIGHT
// ==========================================================
{
    title:"RIGHT",

    keywords:[
        "right",
        "extract last characters",
        "get ending text",
        "last letters",
        "suffix extraction",
        "account number extraction",
        "employee id last digits",
        "text split",
        "data cleaning",
        "text processing",
        "excel text function",
        "advanced excel",
        "mis reporting",
        "dashboard formula",
        "data preparation",
        "automation",
        "business reporting",
        "text analysis",
        "excel formula"
    ],

    aliases:[
        "ending text formula",
        "last character formula",
        "suffix formula"
    ],

    priorityRules:[
        "If query contains first characters => LEFT"
    ],

    questions:[
        "How to extract last characters?",
        "How to get ending text?",
        "How to use RIGHT?"
    ],

    dataset:"account_data",

    category:"Text",

    formula:'=RIGHT(A2,4)',

    syntax:"RIGHT(text,[num_chars])",

    explanation:"Extracts characters from the end of a text string.",

    breakdown:[
        "A2 = Source Text",
        "4 = Characters From Right",
        "Returns Ending Text"
    ],

    output:"4521",

    business:"Used for account numbers, IDs, and suffix extraction.",

    tips:[
        "Useful for last digits extraction.",
        "Combine with LEN for dynamic length."
    ],

    errors:[
        "#VALUE!"
    ],

    alternative:"=TEXTAFTER(A2,\"-\")"
},


// ==========================================================
// MID
// ==========================================================
{
    title:"MID",

    keywords:[
        "mid",
        "middle characters",
        "extract middle text",
        "text from position",
        "substring extraction",
        "employee code split",
        "extract specific characters",
        "text parsing",
        "data cleaning",
        "text processing",
        "excel text function",
        "advanced excel",
        "mis reporting",
        "dashboard formula",
        "data preparation",
        "automation",
        "business reporting",
        "text analysis",
        "excel formula"
    ],

    aliases:[
        "substring formula",
        "middle text extraction",
        "position based text"
    ],

    priorityRules:[
        "If query contains first text => LEFT",
        "If query contains last text => RIGHT"
    ],

    questions:[
        "How to extract middle text?",
        "How to extract characters from position?",
        "How to use MID?"
    ],

    dataset:"employee_id",

    category:"Text",

    formula:'=MID(A2,4,5)',

    syntax:"MID(text,start_num,num_chars)",

    explanation:"Extracts characters from a specific position in text.",

    breakdown:[
        "A2 = Source Text",
        "4 = Starting Position",
        "5 = Characters Required"
    ],

    output:"12345",

    business:"Used for extracting codes from structured IDs.",

    tips:[
        "Position starts from 1.",
        "Useful for fixed-format data."
    ],

    errors:[
        "#VALUE!"
    ],

    alternative:"=TEXTSPLIT(A2,\"-\")"
},


// ==========================================================
// LEN
// ==========================================================
{
    title:"LEN",

    keywords:[
        "len",
        "character count",
        "text length",
        "count characters",
        "string length",
        "data validation",
        "check length",
        "password length",
        "code validation",
        "text analysis",
        "excel text function",
        "advanced excel",
        "mis reporting",
        "dashboard formula",
        "data preparation",
        "automation",
        "business reporting",
        "excel formula"
    ],

    aliases:[
        "length formula",
        "character counter",
        "text count formula"
    ],

    priorityRules:[
        "If query contains remove spaces => TRIM"
    ],

    questions:[
        "How to count characters?",
        "How to check text length?",
        "How to use LEN?"
    ],

    dataset:"customer_data",

    category:"Text",

    formula:"=LEN(A2)",

    syntax:"LEN(text)",

    explanation:"Returns the number of characters in a text string.",

    breakdown:[
        "A2 = Text Value",
        "Counts Letters Numbers Spaces",
        "Returns Length"
    ],

    output:"12",

    business:"Used for validation of IDs, codes, and text fields.",

    tips:[
        "Spaces are counted.",
        "Combine with TRIM for clean data."
    ],

    errors:[
        "#VALUE!"
    ],

    alternative:"=LEN(TRIM(A2))"
},


// ==========================================================
// TRIM
// ==========================================================
{
    title:"TRIM",

    keywords:[
        "trim",
        "remove extra spaces",
        "clean spaces",
        "text cleanup",
        "remove unwanted spaces",
        "data cleaning",
        "imported data cleanup",
        "remove leading spaces",
        "remove trailing spaces",
        "standardize text",
        "excel text function",
        "advanced excel",
        "mis reporting",
        "dashboard formula",
        "data preparation",
        "automation",
        "business reporting",
        "excel formula"
    ],

    aliases:[
        "space remover",
        "text cleaning formula",
        "data cleanup formula"
    ],

    priorityRules:[
        "If query contains clean text => TRIM"
    ],

    questions:[
        "How to remove extra spaces?",
        "How to clean imported data?",
        "How to use TRIM?"
    ],

    dataset:"customer_master",

    category:"Text",

    formula:"=TRIM(A2)",

    syntax:"TRIM(text)",

    explanation:"Removes extra spaces from text except single spaces between words.",

    breakdown:[
        "A2 = Text Value",
        "Removes Leading Spaces",
        "Removes Extra Spaces"
    ],

    output:"John Smith",

    business:"Used for data cleaning before analysis and reporting.",

    tips:[
        "Very useful after CSV imports.",
        "Does not remove non-breaking spaces."
    ],

    errors:[
        "#VALUE!"
    ],

    alternative:"=CLEAN(TRIM(A2))"
},

// ==========================================================
// UPPER
// ==========================================================
{
    title:"UPPER",

    keywords:[
        "upper",
        "uppercase",
        "capital letters",
        "convert to capital",
        "all caps",
        "text capitalization",
        "standardize text",
        "name formatting",
        "data cleaning",
        "text conversion",
        "excel text function",
        "advanced excel",
        "mis reporting",
        "dashboard formula",
        "data preparation",
        "automation",
        "business reporting",
        "excel formula"
    ],

    aliases:[
        "capital text formula",
        "uppercase converter",
        "caps formula"
    ],

    priorityRules:[
        "If query contains small letters => LOWER",
        "If query contains name format => PROPER"
    ],

    questions:[
        "How to convert text to uppercase?",
        "How to make capital letters?",
        "How to use UPPER?"
    ],

    dataset:"employee_master",

    category:"Text",

    formula:"=UPPER(A2)",

    syntax:"UPPER(text)",

    explanation:"Converts all letters in a text string to uppercase.",

    breakdown:[
        "A2 = Source Text",
        "Converts Lowercase Letters",
        "Returns Capital Text"
    ],

    output:"SUNIL KUMAR",

    business:"Used for employee databases and standardized reporting.",

    tips:[
        "Does not change numbers.",
        "Useful before matching text values."
    ],

    errors:[
        "#VALUE!"
    ],

    alternative:"=PROPER(A2)"
},


// ==========================================================
// LOWER
// ==========================================================
{
    title:"LOWER",

    keywords:[
        "lower",
        "lowercase",
        "small letters",
        "convert to lowercase",
        "all small text",
        "text standardization",
        "email formatting",
        "data cleaning",
        "text conversion",
        "excel text function",
        "advanced excel",
        "mis reporting",
        "dashboard formula",
        "data preparation",
        "automation",
        "business reporting",
        "excel formula"
    ],

    aliases:[
        "small letter formula",
        "lowercase converter",
        "email cleanup formula"
    ],

    priorityRules:[
        "If query contains capital letters => UPPER"
    ],

    questions:[
        "How to convert text to lowercase?",
        "How to make small letters?",
        "How to use LOWER?"
    ],

    dataset:"email_data",

    category:"Text",

    formula:"=LOWER(A2)",

    syntax:"LOWER(text)",

    explanation:"Converts all letters in text to lowercase.",

    breakdown:[
        "A2 = Source Text",
        "Changes Capital Letters",
        "Returns Small Letters"
    ],

    output:"sunil.kumar@email.com",

    business:"Used for email cleanup and text standardization.",

    tips:[
        "Useful for duplicate checking.",
        "Common in data preparation."
    ],

    errors:[
        "#VALUE!"
    ],

    alternative:"=UPPER(A2)"
},


// ==========================================================
// PROPER
// ==========================================================
{
    title:"PROPER",

    keywords:[
        "proper",
        "proper case",
        "name format",
        "capitalize first letter",
        "title case",
        "employee name formatting",
        "customer name cleanup",
        "text formatting",
        "data cleaning",
        "standardize names",
        "excel text function",
        "advanced excel",
        "mis reporting",
        "dashboard formula",
        "data preparation",
        "automation",
        "business reporting",
        "excel formula"
    ],

    aliases:[
        "title case formula",
        "name capitalization",
        "first letter capital"
    ],

    priorityRules:[
        "If query contains all capital => UPPER",
        "If query contains all small => LOWER"
    ],

    questions:[
        "How to format names?",
        "How to capitalize first letter?",
        "How to use PROPER?"
    ],

    dataset:"customer_master",

    category:"Text",

    formula:"=PROPER(A2)",

    syntax:"PROPER(text)",

    explanation:"Converts text to proper case where each word starts with a capital letter.",

    breakdown:[
        "A2 = Name Text",
        "Capitalizes First Letter",
        "Converts Remaining Letters"
    ],

    output:"Sunil Kumar",

    business:"Used for employee and customer master data cleaning.",

    tips:[
        "Useful after importing raw data.",
        "Check abbreviations manually."
    ],

    errors:[
        "#VALUE!"
    ],

    alternative:"=TEXTBEFORE(A2,\" \")"
},

// ==========================================================
// CONCAT
// ==========================================================
{
    title:"CONCAT",

    keywords:[
        "concat",
        "combine text",
        "merge cells",
        "join columns",
        "combine values",
        "create full name",
        "text joining",
        "merge employee details",
        "combine address",
        "text function",
        "excel text function",
        "advanced excel",
        "mis reporting",
        "dashboard formula",
        "automation",
        "business reporting",
        "excel formula"
    ],

    aliases:[
        "merge text formula",
        "combine cells formula",
        "join values"
    ],

    priorityRules:[
        "If query contains separator => TEXTJOIN"
    ],

    questions:[
        "How to combine cells?",
        "How to merge text values?",
        "How to use CONCAT?"
    ],

    dataset:"employee_data",

    category:"Text",

    formula:"=CONCAT(A2,B2)",

    syntax:"CONCAT(text1,[text2],...)",

    explanation:"Combines multiple text values into one text string.",

    breakdown:[
        "A2 = First Value",
        "B2 = Second Value",
        "Returns Combined Text"
    ],

    output:"SunilKumar",

    business:"Used for creating IDs, labels, and combined fields.",

    tips:[
        "Does not automatically add spaces.",
        "Use TEXTJOIN for separators."
    ],

    errors:[
        "#VALUE!"
    ],

    alternative:'=A2&" "&B2'
},

// ==========================================================
// SEARCH
// ==========================================================
{
    title:"SEARCH",

    keywords:[
        "search",
        "search text",
        "find word",
        "text lookup",
        "case insensitive search",
        "locate keyword",
        "find phrase",
        "keyword search",
        "contains text",
        "text position",
        "excel text function",
        "data analysis",
        "automation",
        "advanced excel",
        "report preparation",
        "text validation",
        "string search",
        "data cleaning",
        "formula search"
    ],

    aliases:[
        "ignore case search",
        "keyword finder",
        "text locator"
    ],

    priorityRules:[
        "If query contains case insensitive => SEARCH",
        "If query contains exact case => FIND"
    ],

    questions:[
        "How to search text inside cell?",
        "How to find keyword?",
        "How to use SEARCH formula?"
    ],

    dataset:"feedback",

    category:"Text",

    formula:'=SEARCH("urgent",A2)',

    syntax:"SEARCH(find_text,within_text,[start_num])",

    explanation:"Finds the position of text inside another text without case sensitivity.",

    breakdown:[
        "\"urgent\" = Keyword",
        "A2 = Text Cell",
        "Returns Position"
    ],

    output:"12",

    business:"Used for keyword detection, feedback analysis, and text classification.",

    tips:[
        "Not case-sensitive.",
        "Supports wildcards."
    ],

    errors:[
        "#VALUE!"
    ],

    alternative:'=ISNUMBER(SEARCH("urgent",A2))'
},

// ==========================================================
// TODAY
// ==========================================================
{
    title:"TODAY",

    keywords:[
        "today",
        "today date",
        "current date",
        "auto update date",
        "system date",
        "current day",
        "date today",
        "live date",
        "dynamic date",
        "daily report date",
        "age calculation",
        "due date calculation",
        "excel date function",
        "automation",
        "advanced excel",
        "mis reporting",
        "dashboard formula",
        "compliance tracking",
        "date analysis",
        "report preparation"
    ],

    aliases:[
        "current day formula",
        "automatic date",
        "live date formula"
    ],

    priorityRules:[
        "If query contains current time => NOW",
        "If query contains current date only => TODAY"
    ],

    questions:[
        "How to get today's date?",
        "How to insert automatic date?",
        "How to use TODAY formula?"
    ],

    dataset:"compliance",

    category:"Date",

    formula:"=TODAY()",

    syntax:"TODAY()",

    explanation:"Returns the current date automatically.",

    breakdown:[
        "No Arguments Required",
        "Uses System Date",
        "Updates Automatically"
    ],

    output:"28-Jul-2026",

    business:"Used for aging reports, SLA tracking, and daily MIS reports.",

    tips:[
        "Updates when workbook recalculates.",
        "Use VALUE paste if you need a fixed date."
    ],

    errors:[
        "#VALUE!"
    ],

    alternative:"=NOW()"
},


// ==========================================================
// NOW
// ==========================================================
{
    title:"NOW",

    keywords:[
        "now",
        "current date time",
        "current timestamp",
        "live date time",
        "time stamp",
        "record time",
        "transaction time",
        "login time",
        "automation timestamp",
        "real time date",
        "excel date function",
        "excel time function",
        "advanced excel",
        "automation",
        "mis reporting",
        "audit tracking",
        "workflow monitoring",
        "date analysis",
        "business reporting",
        "time tracking"
    ],

    aliases:[
        "timestamp formula",
        "date time formula",
        "live timestamp"
    ],

    priorityRules:[
        "If query contains date and time => NOW",
        "If query contains date only => TODAY"
    ],

    questions:[
        "How to get current date and time?",
        "How to create timestamp?",
        "How to use NOW formula?"
    ],

    dataset:"transaction_log",

    category:"Date",

    formula:"=NOW()",

    syntax:"NOW()",

    explanation:"Returns the current date and time.",

    breakdown:[
        "No Arguments Required",
        "Returns Current Date",
        "Returns Current Time"
    ],

    output:"28-Jul-2026 20:30",

    business:"Used for audit logs, transaction tracking, and workflow monitoring.",

    tips:[
        "Updates automatically.",
        "Useful for timestamp reporting."
    ],

    errors:[
        "#VALUE!"
    ],

    alternative:"=TODAY()"
},


// ==========================================================
// DATE
// ==========================================================
{
    title:"DATE",

    keywords:[
        "date",
        "create date",
        "combine year month day",
        "generate date",
        "build date",
        "date creation",
        "convert values to date",
        "custom date",
        "joining date",
        "expiry date",
        "excel date function",
        "automation",
        "advanced excel",
        "mis reporting",
        "date calculation",
        "business reporting",
        "financial analysis",
        "calendar formula",
        "schedule creation",
        "report preparation"
    ],

    aliases:[
        "date creation formula",
        "make date formula",
        "combine date values"
    ],

    priorityRules:[
        "If query contains extract year => YEAR",
        "If query contains extract month => MONTH"
    ],

    questions:[
        "How to create date from values?",
        "How to combine year month day?",
        "How to use DATE formula?"
    ],

    dataset:"employee_master",

    category:"Date",

    formula:"=DATE(A2,B2,C2)",

    syntax:"DATE(year,month,day)",

    explanation:"Creates a valid Excel date from year, month, and day values.",

    breakdown:[
        "A2 = Year",
        "B2 = Month",
        "C2 = Day"
    ],

    output:"28-Jul-2026",

    business:"Used for joining dates, expiry dates, and schedule creation.",

    tips:[
        "Useful when dates come from separate columns.",
        "Excel stores dates as serial numbers."
    ],

    errors:[
        "#VALUE!",
        "#NUM!"
    ],

    alternative:'=TEXT(DATE(A2,B2,C2),"DD-MMM-YYYY")'
},


// ==========================================================
// YEAR
// ==========================================================
{
    title:"YEAR",

    keywords:[
        "year",
        "extract year",
        "get year from date",
        "joining year",
        "financial year",
        "year analysis",
        "date breakdown",
        "annual report",
        "year wise summary",
        "year calculation",
        "excel date function",
        "automation",
        "advanced excel",
        "mis reporting",
        "dashboard analysis",
        "business reporting",
        "date extraction",
        "calendar analysis",
        "trend analysis",
        "report preparation"
    ],

    aliases:[
        "date year formula",
        "year extraction",
        "annual extraction"
    ],

    priorityRules:[
        "If query contains month => MONTH",
        "If query contains year => YEAR"
    ],

    questions:[
        "How to extract year from date?",
        "How to get joining year?",
        "How to use YEAR formula?"
    ],

    dataset:"employees",

    category:"Date",

    formula:"=YEAR(B2)",

    syntax:"YEAR(serial_number)",

    explanation:"Returns the year from a date.",

    breakdown:[
        "B2 = Date Value",
        "Extracts Year",
        "Returns Four Digit Year"
    ],

    output:"2026",

    business:"Used for yearly analysis, employee reports, and financial reporting.",

    tips:[
        "Combine with MONTH for period analysis.",
        "Works with valid Excel dates."
    ],

    errors:[
        "#VALUE!"
    ],

    alternative:'=TEXT(B2,"YYYY")'
},


// ==========================================================
// MONTH
// ==========================================================
{
    title:"MONTH",

    keywords:[
        "month",
        "extract month",
        "get month from date",
        "joining month",
        "monthly analysis",
        "month wise report",
        "month number",
        "date breakdown",
        "monthly MIS",
        "period analysis",
        "excel date function",
        "automation",
        "advanced excel",
        "mis reporting",
        "dashboard analysis",
        "business reporting",
        "date extraction",
        "calendar analysis",
        "trend analysis",
        "report preparation"
    ],

    aliases:[
        "date month formula",
        "month extraction",
        "monthly calculation"
    ],

    priorityRules:[
        "If query contains year => YEAR",
        "If query contains month => MONTH"
    ],

    questions:[
        "How to extract month from date?",
        "How to get joining month?",
        "How to use MONTH formula?"
    ],

    dataset:"employees",

    category:"Date",

    formula:"=MONTH(B2)",

    syntax:"MONTH(serial_number)",

    explanation:"Returns the month number from a date.",

    breakdown:[
        "B2 = Date Value",
        "Extracts Month",
        "Returns 1 to 12"
    ],

    output:"7",

    business:"Used for monthly MIS, trend analysis, and reporting.",

    tips:[
        "Use TEXT for month names.",
        "Combine with YEAR for monthly grouping."
    ],

    errors:[
        "#VALUE!"
    ],

    alternative:'=TEXT(B2,"MMMM")'
},

// ==========================================================
// IF
// ==========================================================
{
    title:"IF",

    keywords:[
        "if",
        "if condition",
        "logical condition",
        "decision formula",
        "check condition",
        "yes no formula",
        "true false formula",
        "status formula",
        "eligibility formula",
        "classification formula",
        "excel logical function",
        "advanced excel",
        "mis reporting",
        "dashboard formula",
        "automation",
        "business reporting",
        "decision making",
        "excel formula"
    ],

    aliases:[
        "condition formula",
        "logical test formula",
        "decision formula"
    ],

    priorityRules:[
        "If query contains multiple conditions => IFS",
        "If query contains all conditions => AND",
        "If query contains any condition => OR"
    ],

    questions:[
        "How to apply condition?",
        "How to create yes no formula?",
        "How to use IF?"
    ],

    dataset:"employee_status",

    category:"Logical",

    formula:'=IF(B2>=50000,"High","Low")',

    syntax:"IF(logical_test,value_if_true,value_if_false)",

    explanation:"Checks a condition and returns one value if TRUE and another if FALSE.",

    breakdown:[
        "B2>=50000 = Condition",
        "High = TRUE Result",
        "Low = FALSE Result"
    ],

    output:"High",

    business:"Used for employee classification, compliance status, and MIS automation.",

    tips:[
        "Can be nested for multiple decisions.",
        "Use IFS for cleaner multiple conditions."
    ],

    errors:[
        "#VALUE!",
        "#NAME?"
    ],

    alternative:'=IFS(B2>=50000,"High",B2<50000,"Low")'
},


// ==========================================================
// IFS
// ==========================================================
{
    title:"IFS",

    keywords:[
        "ifs",
        "multiple if",
        "multiple conditions",
        "many conditions",
        "multiple decision",
        "grading formula",
        "rating formula",
        "classification formula",
        "nested if alternative",
        "condition based output",
        "excel logical function",
        "advanced excel",
        "mis reporting",
        "dashboard formula",
        "automation",
        "business reporting",
        "decision making",
        "excel formula"
    ],

    aliases:[
        "multiple condition formula",
        "nested if replacement",
        "multi decision formula"
    ],

    priorityRules:[
        "If query contains single condition => IF"
    ],

    questions:[
        "How to use multiple conditions?",
        "How to replace nested IF?",
        "How to use IFS?"
    ],

    dataset:"performance_rating",

    category:"Logical",

    formula:'=IFS(B2>=90,"Excellent",B2>=70,"Good",B2<70,"Average")',

    syntax:"IFS(test1,result1,test2,result2,...)",

    explanation:"Tests multiple conditions and returns the result for the first TRUE condition.",

    breakdown:[
        "B2>=90 = First Condition",
        "Excellent = Result",
        "Checks Conditions Sequentially"
    ],

    output:"Excellent",

    business:"Used for grading, risk categories, and performance ratings.",

    tips:[
        "Cleaner than nested IF.",
        "Available in newer Excel versions."
    ],

    errors:[
        "#N/A",
        "#VALUE!"
    ],

    alternative:'=IF(B2>=90,"Excellent",IF(B2>=70,"Good","Average"))'
},


// ==========================================================
// AND
// ==========================================================
{
    title:"AND",

    keywords:[
        "and",
        "all conditions true",
        "multiple checks",
        "combined conditions",
        "validate all criteria",
        "eligibility check",
        "compliance check",
        "approval formula",
        "logical test",
        "excel logical function",
        "advanced excel",
        "mis reporting",
        "dashboard formula",
        "automation",
        "business reporting",
        "decision making",
        "excel formula"
    ],

    aliases:[
        "all criteria formula",
        "multiple validation formula",
        "combined condition"
    ],

    priorityRules:[
        "If query contains any condition => OR"
    ],

    questions:[
        "How to check multiple conditions together?",
        "How to verify all criteria?",
        "How to use AND?"
    ],

    dataset:"loan_application",

    category:"Logical",

    formula:'=IF(AND(B2>=25,C2="Yes"),"Eligible","Not Eligible")',

    syntax:"AND(logical1,[logical2],...)",

    explanation:"Returns TRUE only when all conditions are TRUE.",

    breakdown:[
        "B2>=25 = Age Check",
        "C2=Yes = Approval Check",
        "Both Must Be True"
    ],

    output:"Eligible",

    business:"Used for compliance checks, approvals, and eligibility rules.",

    tips:[
        "Combine with IF for output.",
        "All conditions must pass."
    ],

    errors:[
        "#VALUE!"
    ],

    alternative:'=IF(B2>=25,"Eligible","Not Eligible")'
},


// ==========================================================
// OR
// ==========================================================
{
    title:"OR",

    keywords:[
        "or",
        "any condition true",
        "either condition",
        "multiple options",
        "alternative criteria",
        "exception check",
        "risk detection",
        "validation formula",
        "logical test",
        "excel logical function",
        "advanced excel",
        "mis reporting",
        "dashboard formula",
        "automation",
        "business reporting",
        "decision making",
        "excel formula"
    ],

    aliases:[
        "either condition formula",
        "any criteria formula",
        "alternative check"
    ],

    priorityRules:[
        "If query contains all conditions => AND"
    ],

    questions:[
        "How to check any condition?",
        "How to use OR logic?",
        "How to combine options?"
    ],

    dataset:"risk_analysis",

    category:"Logical",

    formula:'=IF(OR(B2="High",C2="Critical"),"Alert","Normal")',

    syntax:"OR(logical1,[logical2],...)",

    explanation:"Returns TRUE if at least one condition is TRUE.",

    breakdown:[
        "B2=High = First Check",
        "C2=Critical = Second Check",
        "Any One Can Be True"
    ],

    output:"Alert",

    business:"Used for risk alerts, exception monitoring, and controls testing.",

    tips:[
        "Only one condition needs to match.",
        "Combine with IF for results."
    ],

    errors:[
        "#VALUE!"
    ],

    alternative:'=IF(B2="High","Alert","Normal")'
},


// ==========================================================
// NOT
// ==========================================================
{
    title:"NOT",

    keywords:[
        "not",
        "reverse condition",
        "opposite result",
        "negative logic",
        "exclude condition",
        "check not equal",
        "logical reversal",
        "exception formula",
        "validation formula",
        "excel logical function",
        "advanced excel",
        "mis reporting",
        "dashboard formula",
        "automation",
        "business reporting",
        "decision making",
        "excel formula"
    ],

    aliases:[
        "reverse logic formula",
        "opposite condition",
        "negative test"
    ],

    priorityRules:[
        "If query contains not equal => NOT"
    ],

    questions:[
        "How to reverse condition?",
        "How to check opposite value?",
        "How to use NOT?"
    ],

    dataset:"employee_attendance",

    category:"Logical",

    formula:'=IF(NOT(B2="Absent"),"Present","Absent")',

    syntax:"NOT(logical)",

    explanation:"Reverses the logical value TRUE to FALSE and FALSE to TRUE.",

    breakdown:[
        "B2=Absent = Original Condition",
        "NOT Reverses Result",
        "Returns Opposite Value"
    ],

    output:"Present",

    business:"Used for exception checks and validation rules.",

    tips:[
        "Useful with IF.",
        "Can simplify negative conditions."
    ],

    errors:[
        "#VALUE!"
    ],

    alternative:'=IF(B2<>"Absent","Present","Absent")'
},

// ==========================================================
// IFERROR
// ==========================================================
{
    title:"IFERROR",

    keywords:[
        "iferror",
        "handle errors",
        "remove errors",
        "replace error",
        "avoid formula error",
        "error handling",
        "error free dashboard",
        "hide errors",
        "lookup error handling",
        "formula protection",
        "excel logical function",
        "advanced excel",
        "mis reporting",
        "dashboard formula",
        "automation",
        "business reporting",
        "excel formula"
    ],

    aliases:[
        "error remover formula",
        "safe formula",
        "error control formula"
    ],

    priorityRules:[
        "If query contains #N/A => IFNA"
    ],

    questions:[
        "How to handle formula errors?",
        "How to replace errors?",
        "How to use IFERROR?"
    ],

    dataset:"sales_report",

    category:"Logical",

    formula:'=IFERROR(A2/B2,0)',

    syntax:"IFERROR(value,value_if_error)",

    explanation:"Returns a custom value when a formula results in an error.",

    breakdown:[
        "A2/B2 = Formula Calculation",
        "0 = Error Replacement",
        "Avoids Error Display"
    ],

    output:"0",

    business:"Used for clean dashboards and professional reports.",

    tips:[
        "Handles all Excel errors.",
        "Use IFNA when only #N/A is required."
    ],

    errors:[
        "#VALUE!",
        "#DIV/0!",
        "#N/A"
    ],

    alternative:'=IF(ISERROR(A2/B2),0,A2/B2)'
},


// ==========================================================
// IFNA
// ==========================================================
{
    title:"IFNA",

    keywords:[
        "ifna",
        "n/a error",
        "#N/A handling",
        "lookup error",
        "vlookup error",
        "xlookup error",
        "missing value handling",
        "not available error",
        "lookup protection",
        "excel logical function",
        "advanced excel",
        "mis reporting",
        "dashboard formula",
        "automation",
        "business reporting",
        "excel formula"
    ],

    aliases:[
        "na error remover",
        "lookup error handler",
        "missing value formula"
    ],

    priorityRules:[
        "If query contains all errors => IFERROR"
    ],

    questions:[
        "How to remove #N/A error?",
        "How to handle lookup errors?",
        "How to use IFNA?"
    ],

    dataset:"customer_lookup",

    category:"Logical",

    formula:'=IFNA(XLOOKUP(A2,E:E,F:F),"Not Found")',

    syntax:"IFNA(value,value_if_na)",

    explanation:"Returns a custom value only when the result is #N/A.",

    breakdown:[
        "XLOOKUP = Lookup Formula",
        "Not Found = Replacement",
        "Handles Only #N/A"
    ],

    output:"Not Found",

    business:"Used in lookup reports and reconciliation analysis.",

    tips:[
        "More specific than IFERROR.",
        "Keeps other errors visible."
    ],

    errors:[
        "#N/A"
    ],

    alternative:'=IFERROR(XLOOKUP(A2,E:E,F:F),"Not Found")'
},


// ==========================================================
// SWITCH
// ==========================================================
{
    title:"SWITCH",

    keywords:[
        "switch",
        "multiple value mapping",
        "replace nested if",
        "category mapping",
        "status mapping",
        "rating formula",
        "text classification",
        "multiple choices",
        "condition mapping",
        "excel logical function",
        "advanced excel",
        "mis reporting",
        "dashboard formula",
        "automation",
        "business reporting",
        "excel formula"
    ],

    aliases:[
        "mapping formula",
        "case formula",
        "value based decision"
    ],

    priorityRules:[
        "If query contains multiple IF replacement => SWITCH"
    ],

    questions:[
        "How to map multiple values?",
        "How to replace nested IF?",
        "How to use SWITCH?"
    ],

    dataset:"employee_grade",

    category:"Logical",

    formula:'=SWITCH(A2,"A","Excellent","B","Good","Average")',

    syntax:"SWITCH(expression,value1,result1,...)",

    explanation:"Matches a value against multiple options and returns the related result.",

    breakdown:[
        "A2 = Value To Check",
        "A/B = Matching Options",
        "Returns Corresponding Result"
    ],

    output:"Excellent",

    business:"Used for status mapping, grading, and category conversion.",

    tips:[
        "Cleaner than multiple IF statements.",
        "Best for exact matches."
    ],

    errors:[
        "#N/A",
        "#VALUE!"
    ],

    alternative:'=IF(A2="A","Excellent","Good")'
},


// ==========================================================
// XOR
// ==========================================================
{
    title:"XOR",

    keywords:[
        "xor",
        "exclusive or",
        "only one condition true",
        "one condition check",
        "alternative logic",
        "logical comparison",
        "validation formula",
        "exception logic",
        "excel logical function",
        "advanced excel",
        "mis reporting",
        "dashboard formula",
        "automation",
        "business reporting",
        "excel formula"
    ],

    aliases:[
        "one true condition",
        "exclusive logic",
        "single match formula"
    ],

    priorityRules:[
        "If query contains all true => AND",
        "If query contains any true => OR"
    ],

    questions:[
        "How to check only one condition?",
        "How to use XOR logic?",
        "How to compare alternatives?"
    ],

    dataset:"control_testing",

    category:"Logical",

    formula:'=XOR(B2="Yes",C2="Yes")',

    syntax:"XOR(logical1,[logical2],...)",

    explanation:"Returns TRUE when an odd number of conditions are TRUE.",

    breakdown:[
        "B2 Condition",
        "C2 Condition",
        "Only One Should Match"
    ],

    output:"TRUE",

    business:"Used in validation rules and control testing.",

    tips:[
        "Different from OR.",
        "Two TRUE conditions return FALSE."
    ],

    errors:[
        "#VALUE!"
    ],

    alternative:'=OR(B2="Yes",C2="Yes")'
},

// ==========================================================
// TEXTSPLIT
// ==========================================================
{
    title:"TEXTSPLIT",

    keywords:[
        "textsplit",
        "split text",
        "split by delimiter",
        "split comma separated text",
        "divide text into columns",
        "split values",
        "extract multiple values",
        "dynamic split",
        "excel 365 text function",
        "data cleaning",
        "advanced excel",
        "mis reporting",
        "dashboard formula",
        "automation",
        "excel formula"
    ],

    aliases:[
        "text split formula",
        "delimiter split",
        "split string formula"
    ],

    priorityRules:[
        "If query contains before separator => TEXTBEFORE",
        "If query contains after separator => TEXTAFTER"
    ],

    questions:[
        "How to split text?",
        "How to divide text by comma?",
        "How to use TEXTSPLIT?"
    ],

    dataset:"employee_skills",

    category:"Text",

    formula:'=TEXTSPLIT(A2,",")',

    syntax:"TEXTSPLIT(text,col_delimiter,[row_delimiter])",

    explanation:"Splits text into multiple cells based on a delimiter.",

    breakdown:[
        "A2 = Source Text",
        "\",\" = Separator",
        "Returns Multiple Cells"
    ],

    output:"Excel | Power BI | SQL",

    business:"Used for data cleaning and converting combined values into columns.",

    tips:[
        "Available in Microsoft 365.",
        "Works with comma, space, pipe and other delimiters."
    ],

    errors:[
        "#VALUE!",
        "#CALC!"
    ],

    alternative:'=TEXTBEFORE(A2,",")'
},


// ==========================================================
// TEXTBEFORE
// ==========================================================
{
    title:"TEXTBEFORE",

    keywords:[
        "textbefore",
        "text before delimiter",
        "extract before",
        "get text before comma",
        "extract username",
        "split left side",
        "remove after character",
        "data cleaning",
        "excel 365 text function",
        "advanced excel",
        "mis reporting",
        "dashboard formula",
        "automation",
        "excel formula"
    ],

    aliases:[
        "before separator formula",
        "left extraction advanced",
        "extract first part"
    ],

    priorityRules:[
        "If query contains split complete text => TEXTSPLIT"
    ],

    questions:[
        "How to get text before separator?",
        "How to extract first part?",
        "How to use TEXTBEFORE?"
    ],

    dataset:"email_data",

    category:"Text",

    formula:'=TEXTBEFORE(A2,"@")',

    syntax:"TEXTBEFORE(text,delimiter)",

    explanation:"Returns the text before a specified delimiter.",

    breakdown:[
        "A2 = Source Text",
        "\"@\" = Separator",
        "Returns Before Text"
    ],

    output:"sunil.kumar",

    business:"Used for extracting usernames, codes, and prefixes.",

    tips:[
        "Cleaner alternative to LEFT + FIND.",
        "Available in Excel 365."
    ],

    errors:[
        "#N/A",
        "#VALUE!"
    ],

    alternative:'=LEFT(A2,FIND("@",A2)-1)'
},

// ==========================================================
// MODE.SNGL
// ==========================================================
{
    title:"MODE.SNGL",

    keywords:[
        "mode",
        "most repeated value",
        "most common number",
        "frequent value",
        "highest occurrence",
        "statistical analysis",
        "data pattern",
        "excel statistical function",
        "advanced excel",
        "mis reporting",
        "dashboard formula",
        "excel formula"
    ],

    aliases:[
        "most frequent value",
        "common value formula",
        "repeated number formula"
    ],

    questions:[
        "How to find most repeated value?",
        "How to calculate mode?",
        "How to use MODE?"
    ],

    dataset:"transaction_frequency",

    category:"Statistical",

    formula:"=MODE.SNGL(B2:B100)",

    syntax:"MODE.SNGL(number1,[number2],...)",

    explanation:"Returns the most frequently occurring number.",

    breakdown:[
        "B2:B100 = Data Range",
        "Counts Frequency",
        "Returns Highest Occurrence"
    ],

    output:"5",

    business:"Used for identifying common trends and repeated values.",

    tips:[
        "Works with numeric values only.",
        "Returns first mode if multiple exist."
    ],

    errors:[
        "#N/A",
        "#VALUE!"
    ],

    alternative:"=COUNTIF(B:B,MAX(COUNTIF(B:B,B:B)))"
},


// ==========================================================
// STDEV.S
// ==========================================================
{
    title:"STDEV.S",

    keywords:[
        "stdev",
        "standard deviation",
        "data spread",
        "variation analysis",
        "risk analysis",
        "volatility calculation",
        "statistical deviation",
        "excel statistical function",
        "advanced excel",
        "mis reporting",
        "dashboard formula",
        "excel formula"
    ],

    aliases:[
        "std deviation formula",
        "data variation",
        "risk measurement formula"
    ],

    questions:[
        "How to calculate standard deviation?",
        "How to measure data variation?",
        "How to use STDEV.S?"
    ],

    dataset:"sales_variance",

    category:"Statistical",

    formula:"=STDEV.S(B2:B100)",

    syntax:"STDEV.S(number1,[number2],...)",

    explanation:"Calculates standard deviation based on a sample dataset.",

    breakdown:[
        "B2:B100 = Sample Data",
        "Measures Variation",
        "Returns Deviation Value"
    ],

    output:"1250.45",

    business:"Used for risk analysis, forecasting, and performance variation.",

    tips:[
        "Use STDEV.P for entire population.",
        "Lower value means more consistency."
    ],

    errors:[
        "#DIV/0!",
        "#VALUE!"
    ],

    alternative:"=STDEV.P(B2:B100)"
},


// ==========================================================
// RANK.EQ
// ==========================================================
{
    title:"RANK.EQ",

    keywords:[
        "rank",
        "ranking",
        "position",
        "employee ranking",
        "sales rank",
        "performance ranking",
        "highest value rank",
        "leaderboard",
        "excel statistical function",
        "advanced excel",
        "mis reporting",
        "dashboard formula",
        "excel formula"
    ],

    aliases:[
        "rank formula",
        "position formula",
        "top performer formula"
    ],

    questions:[
        "How to rank values?",
        "How to find position?",
        "How to use RANK.EQ?"
    ],

    dataset:"employee_performance",

    category:"Statistical",

    formula:"=RANK.EQ(B2,$B$2:$B$100,0)",

    syntax:"RANK.EQ(number,ref,[order])",

    explanation:"Returns the rank of a number compared with other values.",

    breakdown:[
        "B2 = Value To Rank",
        "$B$2:$B$100 = Comparison Range",
        "0 = Highest First"
    ],

    output:"1",

    business:"Used for employee ranking, sales ranking, and KPI reports.",

    tips:[
        "Duplicate values get same rank.",
        "Use RANK.AVG for average ranking."
    ],

    errors:[
        "#N/A",
        "#VALUE!"
    ],

    alternative:"=SORT(B2:B100)"
},


// ==========================================================
// PERCENTILE.INC
// ==========================================================
{
    title:"PERCENTILE.INC",

    keywords:[
        "percentile",
        "percentile analysis",
        "data distribution",
        "quartile analysis",
        "top percentage",
        "benchmark analysis",
        "statistical calculation",
        "excel statistical function",
        "advanced excel",
        "mis reporting",
        "dashboard formula",
        "excel formula"
    ],

    aliases:[
        "percentile formula",
        "distribution formula",
        "benchmark formula"
    ],

    questions:[
        "How to calculate percentile?",
        "How to find top percentage?",
        "How to use PERCENTILE.INC?"
    ],

    dataset:"performance_score",

    category:"Statistical",

    formula:"=PERCENTILE.INC(B2:B100,0.90)",

    syntax:"PERCENTILE.INC(array,k)",

    explanation:"Returns the value at a specified percentile of a dataset.",

    breakdown:[
        "B2:B100 = Data Range",
        "0.90 = 90th Percentile",
        "Returns Benchmark Value"
    ],

    output:"95",

    business:"Used for employee benchmarking, risk analysis, and analytics.",

    tips:[
        "k must be between 0 and 1.",
        "Useful for identifying top performers."
    ],

    errors:[
        "#NUM!",
        "#VALUE!"
    ],

    alternative:"=QUARTILE.INC(B2:B100,4)"
},

// ==========================================================
// TEXTAFTER
// ==========================================================
{
    title:"TEXTAFTER",

    keywords:[
        "textafter",
        "text after delimiter",
        "extract after",
        "get text after comma",
        "extract domain",
        "split right side",
        "remove before character",
        "data cleaning",
        "excel 365 text function",
        "advanced excel",
        "mis reporting",
        "dashboard formula",
        "automation",
        "excel formula"
    ],

    aliases:[
        "after separator formula",
        "right extraction advanced",
        "extract second part"
    ],

    priorityRules:[
        "If query contains before separator => TEXTBEFORE"
    ],

    questions:[
        "How to get text after separator?",
        "How to extract remaining text?",
        "How to use TEXTAFTER?"
    ],

    dataset:"email_data",

    category:"Text",

    formula:'=TEXTAFTER(A2,"@")',

    syntax:"TEXTAFTER(text,delimiter)",

    explanation:"Returns the text after a specified delimiter.",

    breakdown:[
        "A2 = Source Text",
        "\"@\" = Separator",
        "Returns After Text"
    ],

    output:"gmail.com",

    business:"Used for extracting domains, categories, and suffix values.",

    tips:[
        "Cleaner alternative to RIGHT + FIND.",
        "Works with multiple delimiters."
    ],

    errors:[
        "#N/A",
        "#VALUE!"
    ],

    alternative:'=RIGHT(A2,LEN(A2)-FIND("@",A2))'
},


// ==========================================================
// SUBSTITUTE
// ==========================================================
{
    title:"SUBSTITUTE",

    keywords:[
        "substitute",
        "replace text",
        "change word",
        "replace specific text",
        "remove characters",
        "text cleanup",
        "standardize values",
        "data correction",
        "replace old value",
        "excel text function",
        "advanced excel",
        "mis reporting",
        "dashboard formula",
        "excel formula"
    ],

    aliases:[
        "text replacement formula",
        "change text formula",
        "find replace formula"
    ],

    priorityRules:[
        "If query contains position replacement => REPLACE"
    ],

    questions:[
        "How to replace text?",
        "How to change specific word?",
        "How to use SUBSTITUTE?"
    ],

    dataset:"customer_data",

    category:"Text",

    formula:'=SUBSTITUTE(A2,"Ltd","Limited")',

    syntax:"SUBSTITUTE(text,old_text,new_text,[instance_num])",

    explanation:"Replaces existing text with new text.",

    breakdown:[
        "A2 = Original Text",
        "Ltd = Old Text",
        "Limited = New Text"
    ],

    output:"ABC Limited",

    business:"Used for data standardization and cleanup.",

    tips:[
        "Can replace specific occurrence.",
        "Case sensitive function."
    ],

    errors:[
        "#VALUE!"
    ],

    alternative:'=REPLACE(A2,1,3,"New")'
},


// ==========================================================
// REPLACE
// ==========================================================
{
    title:"REPLACE",

    keywords:[
        "replace",
        "replace characters",
        "replace by position",
        "change fixed length text",
        "modify text position",
        "mask data",
        "hide characters",
        "text modification",
        "excel text function",
        "advanced excel",
        "mis reporting",
        "dashboard formula",
        "automation",
        "excel formula"
    ],

    aliases:[
        "position replace formula",
        "character replacement",
        "fixed text replacement"
    ],

    priorityRules:[
        "If query contains specific word => SUBSTITUTE"
    ],

    questions:[
        "How to replace characters by position?",
        "How to modify text?",
        "How to use REPLACE?"
    ],

    dataset:"account_number",

    category:"Text",

    formula:'=REPLACE(A2,1,4,"XXXX")',

    syntax:"REPLACE(old_text,start_num,num_chars,new_text)",

    explanation:"Replaces characters based on their position in text.",

    breakdown:[
        "A2 = Original Text",
        "1 = Starting Position",
        "4 = Characters To Replace"
    ],

    output:"XXXX5678",

    business:"Used for masking account numbers and sensitive data.",

    tips:[
        "Position based replacement.",
        "Use SUBSTITUTE for word replacement."
    ],

    errors:[
        "#VALUE!"
    ],

    alternative:'=SUBSTITUTE(A2,"1234","XXXX")'
},

// ==========================================================
// TRUE
// ==========================================================
{
    title:"TRUE",

    keywords:[
        "true",
        "true value",
        "logical true",
        "return true",
        "boolean true",
        "condition result",
        "logical function",
        "excel logical function",
        "advanced excel",
        "mis reporting",
        "dashboard formula",
        "automation",
        "excel formula"
    ],

    aliases:[
        "boolean true",
        "yes value"
    ],

    questions:[
        "How to return TRUE value?",
        "How to use TRUE?"
    ],

    dataset:"validation",

    category:"Logical",

    formula:"=TRUE()",

    syntax:"TRUE()",

    explanation:"Returns the logical value TRUE.",

    breakdown:[
        "No Arguments Required",
        "Returns Boolean Value",
        "Used In Logical Tests"
    ],

    output:"TRUE",

    business:"Used in testing and logical calculations.",

    tips:[
        "TRUE equals 1 in calculations."
    ],

    errors:[
        "#VALUE!"
    ],

    alternative:"=1=1"
},


// ==========================================================
// FALSE
// ==========================================================
{
    title:"FALSE",

    keywords:[
        "false",
        "false value",
        "logical false",
        "return false",
        "boolean false",
        "condition result",
        "logical function",
        "excel logical function",
        "advanced excel",
        "mis reporting",
        "dashboard formula",
        "automation",
        "excel formula"
    ],

    aliases:[
        "boolean false",
        "no value"
    ],

    questions:[
        "How to return FALSE value?",
        "How to use FALSE?"
    ],

    dataset:"validation",

    category:"Logical",

    formula:"=FALSE()",

    syntax:"FALSE()",

    explanation:"Returns the logical value FALSE.",

    breakdown:[
        "No Arguments Required",
        "Returns Boolean Value",
        "Used In Logical Tests"
    ],

    output:"FALSE",

    business:"Used for validation and logical checks.",

    tips:[
        "FALSE equals 0 in calculations."
    ],

    errors:[
        "#VALUE!"
    ],

    alternative:"=1=2"
},

// ==========================================================
// DAVERAGE
// ==========================================================
{
    title:"DAVERAGE",

    keywords:[
        "daverage",
        "database average",
        "average with criteria",
        "average from database",
        "criteria based average",
        "database function",
        "excel database formula",
        "advanced excel",
        "mis reporting",
        "dashboard formula",
        "excel formula"
    ],

    aliases:[
        "database average formula",
        "criteria average database",
        "conditional database average"
    ],

    priorityRules:[
        "If query contains count => DCOUNT",
        "If query contains total => DSUM"
    ],

    questions:[
        "How to average database records?",
        "How to calculate criteria based average?",
        "How to use DAVERAGE?"
    ],

    dataset:"employee_database",

    category:"Database",

    formula:'=DAVERAGE(A1:D100,"Salary",F1:F2)',

    syntax:"DAVERAGE(database,field,criteria)",

    explanation:"Calculates the average of values in a database field that meet specified criteria.",

    breakdown:[
        "A1:D100 = Database Range",
        "Salary = Field Name",
        "F1:F2 = Criteria Range"
    ],

    output:"65000",

    business:"Used for employee analysis and criteria-based MIS reports.",

    tips:[
        "Requires column headers.",
        "Criteria must match database headers."
    ],

    errors:[
        "#VALUE!",
        "#NUM!"
    ],

    alternative:"=AVERAGEIFS()"
},


// ==========================================================
// DCOUNT
// ==========================================================
{
    title:"DCOUNT",

    keywords:[
        "dcount",
        "database count",
        "count numbers database",
        "count records with criteria",
        "numeric count",
        "database function",
        "excel database formula",
        "advanced excel",
        "mis reporting",
        "dashboard formula",
        "excel formula"
    ],

    aliases:[
        "database counting formula",
        "criteria number count",
        "conditional database count"
    ],

    priorityRules:[
        "If query contains count all records => DCOUNTA"
    ],

    questions:[
        "How to count database numbers?",
        "How to count records with criteria?",
        "How to use DCOUNT?"
    ],

    dataset:"transaction_database",

    category:"Database",

    formula:'=DCOUNT(A1:E100,"Amount",G1:G2)',

    syntax:"DCOUNT(database,field,criteria)",

    explanation:"Counts numeric values in a database field matching criteria.",

    breakdown:[
        "A1:E100 = Database",
        "Amount = Numeric Field",
        "G1:G2 = Criteria"
    ],

    output:"250",

    business:"Used for transaction analysis and numeric record counting.",

    tips:[
        "Counts only numbers.",
        "Use DCOUNTA for all non-empty records."
    ],

    errors:[
        "#VALUE!",
        "#NUM!"
    ],

    alternative:"=COUNTIFS()"
},


// ==========================================================
// DCOUNTA
// ==========================================================
{
    title:"DCOUNTA",

    keywords:[
        "dcounta",
        "database count all",
        "count records",
        "count non blank",
        "database records count",
        "data completeness",
        "database function",
        "excel database formula",
        "advanced excel",
        "mis reporting",
        "dashboard formula",
        "excel formula"
    ],

    aliases:[
        "database record count",
        "non blank database count",
        "record counting formula"
    ],

    questions:[
        "How to count database records?",
        "How to count non blank values?",
        "How to use DCOUNTA?"
    ],

    dataset:"customer_database",

    category:"Database",

    formula:'=DCOUNTA(A1:E100,"Customer_ID",G1:G2)',

    syntax:"DCOUNTA(database,field,criteria)",

    explanation:"Counts non-empty cells in a database field based on criteria.",

    breakdown:[
        "A1:E100 = Database Range",
        "Customer_ID = Field",
        "G1:G2 = Criteria"
    ],

    output:"1500",

    business:"Used for customer records, employee lists, and data quality checks.",

    tips:[
        "Counts text and numbers.",
        "Useful for completeness reports."
    ],

    errors:[
        "#VALUE!"
    ],

    alternative:"=COUNTAIFS()"
},


// ==========================================================
// DGET
// ==========================================================
{
    title:"DGET",

    keywords:[
        "dget",
        "get single record",
        "extract matching value",
        "database lookup",
        "single result lookup",
        "criteria lookup",
        "database function",
        "excel database formula",
        "advanced excel",
        "mis reporting",
        "dashboard formula",
        "excel formula"
    ],

    aliases:[
        "database lookup formula",
        "single record extraction",
        "criteria based lookup"
    ],

    questions:[
        "How to extract one database value?",
        "How to get matching record?",
        "How to use DGET?"
    ],

    dataset:"employee_database",

    category:"Database",

    formula:'=DGET(A1:F100,"Email",H1:H2)',

    syntax:"DGET(database,field,criteria)",

    explanation:"Returns a single value from a database that matches criteria.",

    breakdown:[
        "A1:F100 = Database",
        "Email = Required Field",
        "H1:H2 = Search Criteria"
    ],

    output:"sunil.kumar@email.com",

    business:"Used for extracting unique records from structured databases.",

    tips:[
        "Requires exactly one matching record.",
        "Multiple matches return an error."
    ],

    errors:[
        "#NUM!",
        "#VALUE!"
    ],

    alternative:"=XLOOKUP()"
},


// ==========================================================
// DSUM
// ==========================================================
{
    title:"DSUM",

    keywords:[
        "dsum",
        "database sum",
        "sum with criteria",
        "criteria based total",
        "database total",
        "conditional database calculation",
        "database function",
        "excel database formula",
        "advanced excel",
        "mis reporting",
        "dashboard formula",
        "excel formula"
    ],

    aliases:[
        "database total formula",
        "criteria sum database",
        "conditional database sum"
    ],

    priorityRules:[
        "If query contains average => DAVERAGE"
    ],

    questions:[
        "How to sum database values?",
        "How to calculate criteria total?",
        "How to use DSUM?"
    ],

    dataset:"sales_database",

    category:"Database",

    formula:'=DSUM(A1:F100,"Sales",H1:H2)',

    syntax:"DSUM(database,field,criteria)",

    explanation:"Adds values from a database field that meet specified criteria.",

    breakdown:[
        "A1:F100 = Database",
        "Sales = Sum Field",
        "H1:H2 = Criteria"
    ],

    output:"450000",

    business:"Used for sales MIS, financial reports, and audit analysis.",

    tips:[
        "Requires database headers.",
        "Useful for structured reporting."
    ],

    errors:[
        "#VALUE!",
        "#NUM!"
    ],

    alternative:"=SUMIFS()"
},

// ==========================================================
// CONVERT
// ==========================================================
{
    title:"CONVERT",

    keywords:[
        "convert",
        "unit conversion",
        "change measurement",
        "convert kg to pound",
        "convert km to miles",
        "engineering conversion",
        "measurement conversion",
        "excel engineering function",
        "advanced excel",
        "data transformation",
        "excel formula"
    ],

    aliases:[
        "unit converter formula",
        "measurement conversion",
        "conversion formula"
    ],

    questions:[
        "How to convert units?",
        "How to change measurement?",
        "How to use CONVERT?"
    ],

    dataset:"measurement_data",

    category:"Engineering",

    formula:'=CONVERT(A2,"kg","lbm")',

    syntax:"CONVERT(number,from_unit,to_unit)",

    explanation:"Converts a number from one measurement unit to another.",

    breakdown:[
        "A2 = Value",
        "kg = Original Unit",
        "lbm = Target Unit"
    ],

    output:"220.46",

    business:"Used in engineering reports and operational calculations.",

    tips:[
        "Supports many measurement categories.",
        "Unit names must be valid Excel codes."
    ],

    errors:[
        "#N/A",
        "#VALUE!"
    ],

    alternative:"Manual conversion formula"
},


// ==========================================================
// DECIMAL
// ==========================================================
{
    title:"DECIMAL",

    keywords:[
        "decimal",
        "convert text number",
        "base number to decimal",
        "binary conversion",
        "hex conversion",
        "number system conversion",
        "excel engineering function",
        "advanced excel",
        "data conversion",
        "excel formula"
    ],

    aliases:[
        "base to decimal formula",
        "number system converter"
    ],

    questions:[
        "How to convert base number to decimal?",
        "How to use DECIMAL?"
    ],

    dataset:"number_system",

    category:"Engineering",

    formula:'=DECIMAL("1010",2)',

    syntax:"DECIMAL(text,radix)",

    explanation:"Converts a number from another base system into decimal.",

    breakdown:[
        "1010 = Binary Number",
        "2 = Base",
        "Returns Decimal Value"
    ],

    output:"10",

    business:"Used in technical calculations and system conversions.",

    tips:[
        "Supports bases from 2 to 36.",
        "Input is treated as text."
    ],

    errors:[
        "#NUM!",
        "#VALUE!"
    ],

    alternative:"=VALUE(A2)"
},


// ==========================================================
// BASE
// ==========================================================
{
    title:"BASE",

    keywords:[
        "base",
        "decimal to binary",
        "convert decimal",
        "number system",
        "binary conversion",
        "hex conversion",
        "engineering function",
        "excel formula"
    ],

    aliases:[
        "decimal converter",
        "number base conversion"
    ],

    questions:[
        "How to convert decimal to binary?",
        "How to change number base?",
        "How to use BASE?"
    ],

    dataset:"technical_values",

    category:"Engineering",

    formula:"=BASE(10,2)",

    syntax:"BASE(number,radix,[min_length])",

    explanation:"Converts a decimal number into another base system.",

    breakdown:[
        "10 = Decimal Number",
        "2 = Target Base",
        "Returns Binary Value"
    ],

    output:"1010",

    business:"Used in programming, technical analysis, and system reports.",

    tips:[
        "Useful for binary and hexadecimal conversion.",
        "Supports bases 2 to 36."
    ],

    errors:[
        "#NUM!",
        "#VALUE!"
    ],

    alternative:'=DECIMAL(A2,2)'
},


// ==========================================================
// ROMAN
// ==========================================================
{
    title:"ROMAN",

    keywords:[
        "roman",
        "number to roman",
        "convert to roman numerals",
        "roman number formula",
        "text conversion",
        "excel math function",
        "excel formula"
    ],

    aliases:[
        "roman numeral converter",
        "number format converter"
    ],

    questions:[
        "How to convert number to roman?",
        "How to use ROMAN?"
    ],

    dataset:"number_format",

    category:"Engineering",

    formula:"=ROMAN(2026)",

    syntax:"ROMAN(number,[form])",

    explanation:"Converts an Arabic number into Roman numerals.",

    breakdown:[
        "2026 = Number",
        "Converts Numeric Value",
        "Returns Roman Text"
    ],

    output:"MMXXVI",

    business:"Used for numbering, documentation, and formatting.",

    tips:[
        "Returns text output.",
        "Useful for reports and headings."
    ],

    errors:[
        "#VALUE!",
        "#NUM!"
    ],

    alternative:"Manual Roman conversion"
},


// ==========================================================
// ARABIC
// ==========================================================
{
    title:"ARABIC",

    keywords:[
        "arabic",
        "roman to number",
        "convert roman numeral",
        "roman conversion",
        "number extraction",
        "text conversion",
        "excel math function",
        "excel formula"
    ],

    aliases:[
        "roman converter reverse",
        "roman number to decimal"
    ],

    questions:[
        "How to convert roman number to normal number?",
        "How to use ARABIC?"
    ],

    dataset:"document_numbering",

    category:"Engineering",

    formula:'=ARABIC("MMXXVI")',

    syntax:"ARABIC(text)",

    explanation:"Converts Roman numerals into Arabic numbers.",

    breakdown:[
        "MMXXVI = Roman Text",
        "Converts To Number",
        "Returns Decimal Value"
    ],

    output:"2026",

    business:"Used for document numbering and data conversion.",

    tips:[
        "Input should be valid Roman numeral text.",
        "Available in modern Excel versions."
    ],

    errors:[
        "#VALUE!"
    ],

    alternative:"=VALUE(A2)"
},

// ==========================================================
// RATE
// ==========================================================
{
    title:"RATE",

    keywords:[
        "rate",
        "interest rate",
        "find interest rate",
        "loan interest calculation",
        "financial rate",
        "investment return rate",
        "banking calculation",
        "excel financial function",
        "advanced excel",
        "treasury analysis",
        "excel formula"
    ],

    aliases:[
        "interest percentage formula",
        "loan rate formula",
        "return rate calculation"
    ],

    questions:[
        "How to calculate interest rate?",
        "How to find loan rate?",
        "How to use RATE?"
    ],

    dataset:"loan_details",

    category:"Financial",

    formula:"=RATE(60,-1000,50000)",

    syntax:"RATE(nper,pmt,pv,[fv],[type],[guess])",

    explanation:"Calculates the interest rate per period for a loan or investment.",

    breakdown:[
        "60 = Number of Periods",
        "-1000 = Payment Amount",
        "50000 = Present Value"
    ],

    output:"1.5%",

    business:"Used for loan analysis and financial modelling.",

    tips:[
        "Returns rate per payment period.",
        "Combine with annual conversion when required."
    ],

    errors:[
        "#NUM!",
        "#VALUE!"
    ],

    alternative:"=IRR()"
},


// ==========================================================
// NPER
// ==========================================================
{
    title:"NPER",

    keywords:[
        "nper",
        "number of periods",
        "loan duration",
        "payment periods",
        "installment count",
        "financial calculation",
        "loan analysis",
        "excel financial function",
        "advanced excel",
        "excel formula"
    ],

    aliases:[
        "loan tenure formula",
        "number installment formula",
        "payment period formula"
    ],

    questions:[
        "How to calculate loan period?",
        "How many payments required?",
        "How to use NPER?"
    ],

    dataset:"loan_schedule",

    category:"Financial",

    formula:"=NPER(10%/12,-1000,50000)",

    syntax:"NPER(rate,pmt,pv,[fv],[type])",

    explanation:"Calculates the number of payment periods required for a loan or investment.",

    breakdown:[
        "10%/12 = Monthly Interest Rate",
        "-1000 = Payment",
        "50000 = Loan Amount"
    ],

    output:"58 Months",

    business:"Used for EMI and loan tenure calculations.",

    tips:[
        "Useful for repayment planning.",
        "Works with regular payments."
    ],

    errors:[
        "#NUM!",
        "#VALUE!"
    ],

    alternative:"=PMT()"
},


// ==========================================================
// EFFECT
// ==========================================================
{
    title:"EFFECT",

    keywords:[
        "effect",
        "effective interest rate",
        "annual effective rate",
        "compounded interest",
        "interest conversion",
        "financial analysis",
        "excel financial function",
        "advanced excel",
        "excel formula"
    ],

    aliases:[
        "effective rate formula",
        "compound interest rate",
        "annual interest conversion"
    ],

    questions:[
        "How to calculate effective rate?",
        "How to convert nominal rate?",
        "How to use EFFECT?"
    ],

    dataset:"interest_rate",

    category:"Financial",

    formula:"=EFFECT(12%,12)",

    syntax:"EFFECT(nominal_rate,npery)",

    explanation:"Returns the effective annual interest rate based on nominal rate and compounding periods.",

    breakdown:[
        "12% = Nominal Rate",
        "12 = Compounding Periods",
        "Returns Effective Rate"
    ],

    output:"12.68%",

    business:"Used in banking and investment comparison.",

    tips:[
        "Higher compounding frequency increases effective rate.",
        "Use NOMINAL for reverse calculation."
    ],

    errors:[
        "#NUM!",
        "#VALUE!"
    ],

    alternative:"=NOMINAL()"
},


// ==========================================================
// NOMINAL
// ==========================================================
{
    title:"NOMINAL",

    keywords:[
        "nominal",
        "nominal interest rate",
        "convert effective rate",
        "annual nominal rate",
        "interest conversion",
        "financial analysis",
        "excel financial function",
        "advanced excel",
        "excel formula"
    ],

    aliases:[
        "nominal rate formula",
        "effective to nominal conversion"
    ],

    questions:[
        "How to calculate nominal rate?",
        "How to convert effective interest?",
        "How to use NOMINAL?"
    ],

    dataset:"investment_data",

    category:"Financial",

    formula:"=NOMINAL(12.68%,12)",

    syntax:"NOMINAL(effect_rate,npery)",

    explanation:"Converts effective annual interest rate into nominal annual rate.",

    breakdown:[
        "12.68% = Effective Rate",
        "12 = Compounding Periods",
        "Returns Nominal Rate"
    ],

    output:"12%",

    business:"Used for comparing financial products and interest rates.",

    tips:[
        "Opposite of EFFECT function.",
        "Compounding frequency is required."
    ],

    errors:[
        "#NUM!",
        "#VALUE!"
    ],

    alternative:"=EFFECT()"
},


// ==========================================================
// IPMT
// ==========================================================
{
    title:"IPMT",

    keywords:[
        "ipmt",
        "interest payment",
        "loan interest calculation",
        "emi interest part",
        "period interest",
        "loan schedule",
        "financial function",
        "banking calculation",
        "advanced excel",
        "excel formula"
    ],

    aliases:[
        "interest installment formula",
        "loan interest component",
        "emi interest formula"
    ],

    questions:[
        "How to calculate interest payment?",
        "How to find EMI interest portion?",
        "How to use IPMT?"
    ],

    dataset:"loan_repayment",

    category:"Financial",

    formula:"=IPMT(10%/12,1,60,50000)",

    syntax:"IPMT(rate,per,nper,pv,[fv],[type])",

    explanation:"Calculates the interest portion of a loan payment for a specific period.",

    breakdown:[
        "10%/12 = Monthly Rate",
        "1 = Payment Period",
        "50000 = Loan Amount"
    ],

    output:"-$416.67",

    business:"Used in EMI schedules and loan repayment analysis.",

    tips:[
        "Returns only interest component.",
        "Use PPMT for principal component."
    ],

    errors:[
        "#NUM!",
        "#VALUE!"
    ],

    alternative:"=PPMT()"
},

// ==========================================================
// RATE
// ==========================================================
{
    title:"RATE",

    keywords:[
        "rate",
        "interest rate",
        "find interest rate",
        "loan interest calculation",
        "financial rate",
        "investment return rate",
        "banking calculation",
        "excel financial function",
        "advanced excel",
        "treasury analysis",
        "excel formula"
    ],

    aliases:[
        "interest percentage formula",
        "loan rate formula",
        "return rate calculation"
    ],

    questions:[
        "How to calculate interest rate?",
        "How to find loan rate?",
        "How to use RATE?"
    ],

    dataset:"loan_details",

    category:"Financial",

    formula:"=RATE(60,-1000,50000)",

    syntax:"RATE(nper,pmt,pv,[fv],[type],[guess])",

    explanation:"Calculates the interest rate per period for a loan or investment.",

    breakdown:[
        "60 = Number of Periods",
        "-1000 = Payment Amount",
        "50000 = Present Value"
    ],

    output:"1.5%",

    business:"Used for loan analysis and financial modelling.",

    tips:[
        "Returns rate per payment period.",
        "Combine with annual conversion when required."
    ],

    errors:[
        "#NUM!",
        "#VALUE!"
    ],

    alternative:"=IRR()"
},


// ==========================================================
// NPER
// ==========================================================
{
    title:"NPER",

    keywords:[
        "nper",
        "number of periods",
        "loan duration",
        "payment periods",
        "installment count",
        "financial calculation",
        "loan analysis",
        "excel financial function",
        "advanced excel",
        "excel formula"
    ],

    aliases:[
        "loan tenure formula",
        "number installment formula",
        "payment period formula"
    ],

    questions:[
        "How to calculate loan period?",
        "How many payments required?",
        "How to use NPER?"
    ],

    dataset:"loan_schedule",

    category:"Financial",

    formula:"=NPER(10%/12,-1000,50000)",

    syntax:"NPER(rate,pmt,pv,[fv],[type])",

    explanation:"Calculates the number of payment periods required for a loan or investment.",

    breakdown:[
        "10%/12 = Monthly Interest Rate",
        "-1000 = Payment",
        "50000 = Loan Amount"
    ],

    output:"58 Months",

    business:"Used for EMI and loan tenure calculations.",

    tips:[
        "Useful for repayment planning.",
        "Works with regular payments."
    ],

    errors:[
        "#NUM!",
        "#VALUE!"
    ],

    alternative:"=PMT()"
},


// ==========================================================
// EFFECT
// ==========================================================
{
    title:"EFFECT",

    keywords:[
        "effect",
        "effective interest rate",
        "annual effective rate",
        "compounded interest",
        "interest conversion",
        "financial analysis",
        "excel financial function",
        "advanced excel",
        "excel formula"
    ],

    aliases:[
        "effective rate formula",
        "compound interest rate",
        "annual interest conversion"
    ],

    questions:[
        "How to calculate effective rate?",
        "How to convert nominal rate?",
        "How to use EFFECT?"
    ],

    dataset:"interest_rate",

    category:"Financial",

    formula:"=EFFECT(12%,12)",

    syntax:"EFFECT(nominal_rate,npery)",

    explanation:"Returns the effective annual interest rate based on nominal rate and compounding periods.",

    breakdown:[
        "12% = Nominal Rate",
        "12 = Compounding Periods",
        "Returns Effective Rate"
    ],

    output:"12.68%",

    business:"Used in banking and investment comparison.",

    tips:[
        "Higher compounding frequency increases effective rate.",
        "Use NOMINAL for reverse calculation."
    ],

    errors:[
        "#NUM!",
        "#VALUE!"
    ],

    alternative:"=NOMINAL()"
},


// ==========================================================
// NOMINAL
// ==========================================================
{
    title:"NOMINAL",

    keywords:[
        "nominal",
        "nominal interest rate",
        "convert effective rate",
        "annual nominal rate",
        "interest conversion",
        "financial analysis",
        "excel financial function",
        "advanced excel",
        "excel formula"
    ],

    aliases:[
        "nominal rate formula",
        "effective to nominal conversion"
    ],

    questions:[
        "How to calculate nominal rate?",
        "How to convert effective interest?",
        "How to use NOMINAL?"
    ],

    dataset:"investment_data",

    category:"Financial",

    formula:"=NOMINAL(12.68%,12)",

    syntax:"NOMINAL(effect_rate,npery)",

    explanation:"Converts effective annual interest rate into nominal annual rate.",

    breakdown:[
        "12.68% = Effective Rate",
        "12 = Compounding Periods",
        "Returns Nominal Rate"
    ],

    output:"12%",

    business:"Used for comparing financial products and interest rates.",

    tips:[
        "Opposite of EFFECT function.",
        "Compounding frequency is required."
    ],

    errors:[
        "#NUM!",
        "#VALUE!"
    ],

    alternative:"=EFFECT()"
},


// ==========================================================
// IPMT
// ==========================================================
{
    title:"IPMT",

    keywords:[
        "ipmt",
        "interest payment",
        "loan interest calculation",
        "emi interest part",
        "period interest",
        "loan schedule",
        "financial function",
        "banking calculation",
        "advanced excel",
        "excel formula"
    ],

    aliases:[
        "interest installment formula",
        "loan interest component",
        "emi interest formula"
    ],

    questions:[
        "How to calculate interest payment?",
        "How to find EMI interest portion?",
        "How to use IPMT?"
    ],

    dataset:"loan_repayment",

    category:"Financial",

    formula:"=IPMT(10%/12,1,60,50000)",

    syntax:"IPMT(rate,per,nper,pv,[fv],[type])",

    explanation:"Calculates the interest portion of a loan payment for a specific period.",

    breakdown:[
        "10%/12 = Monthly Rate",
        "1 = Payment Period",
        "50000 = Loan Amount"
    ],

    output:"-$416.67",

    business:"Used in EMI schedules and loan repayment analysis.",

    tips:[
        "Returns only interest component.",
        "Use PPMT for principal component."
    ],

    errors:[
        "#NUM!",
        "#VALUE!"
    ],

    alternative:"=PPMT()"
},

// ==========================================================
// PPMT
// ==========================================================
{
    title:"PPMT",

    keywords:[
        "ppmt",
        "principal payment",
        "loan principal calculation",
        "emi principal part",
        "principal installment",
        "loan schedule",
        "financial function",
        "banking calculation",
        "advanced excel",
        "excel formula"
    ],

    aliases:[
        "principal component formula",
        "emi principal formula",
        "loan repayment principal"
    ],

    priorityRules:[
        "If query contains interest part => IPMT"
    ],

    questions:[
        "How to calculate principal payment?",
        "How to find EMI principal portion?",
        "How to use PPMT?"
    ],

    dataset:"loan_repayment",

    category:"Financial",

    formula:"=PPMT(10%/12,1,60,50000)",

    syntax:"PPMT(rate,per,nper,pv,[fv],[type])",

    explanation:"Calculates the principal portion of a loan payment for a specific period.",

    breakdown:[
        "10%/12 = Monthly Interest Rate",
        "1 = Payment Period",
        "50000 = Loan Amount"
    ],

    output:"-$833.33",

    business:"Used for loan amortization schedules and EMI analysis.",

    tips:[
        "Returns only principal amount.",
        "Use IPMT for interest portion."
    ],

    errors:[
        "#NUM!",
        "#VALUE!"
    ],

    alternative:"=IPMT()"
},


// ==========================================================
// CUMIPMT
// ==========================================================
{
    title:"CUMIPMT",

    keywords:[
        "cumipmt",
        "total interest paid",
        "cumulative interest",
        "loan interest total",
        "emi interest calculation",
        "loan analysis",
        "financial function",
        "banking calculation",
        "advanced excel",
        "excel formula"
    ],

    aliases:[
        "total loan interest formula",
        "interest accumulation formula",
        "loan cost formula"
    ],

    questions:[
        "How to calculate total interest paid?",
        "How to find cumulative interest?",
        "How to use CUMIPMT?"
    ],

    dataset:"loan_schedule",

    category:"Financial",

    formula:"=CUMIPMT(10%/12,60,50000,1,60,0)",

    syntax:"CUMIPMT(rate,nper,pv,start_period,end_period,type)",

    explanation:"Calculates total interest paid between two loan periods.",

    breakdown:[
        "Rate = Loan Interest",
        "Periods = Loan Tenure",
        "Returns Total Interest"
    ],

    output:"-$16200",

    business:"Used for loan cost analysis and banking reports.",

    tips:[
        "Useful for complete loan period analysis.",
        "Returns cumulative interest."
    ],

    errors:[
        "#NUM!",
        "#VALUE!"
    ],

    alternative:"=SUM(IPMT())"
},


// ==========================================================
// CUMPRINC
// ==========================================================
{
    title:"CUMPRINC",

    keywords:[
        "cumprinc",
        "total principal paid",
        "cumulative principal",
        "loan principal calculation",
        "emi principal total",
        "repayment analysis",
        "financial function",
        "advanced excel",
        "excel formula"
    ],

    aliases:[
        "principal accumulation formula",
        "loan repayment total",
        "principal paid formula"
    ],

    questions:[
        "How to calculate total principal paid?",
        "How to track loan principal?",
        "How to use CUMPRINC?"
    ],

    dataset:"loan_schedule",

    category:"Financial",

    formula:"=CUMPRINC(10%/12,60,50000,1,60,0)",

    syntax:"CUMPRINC(rate,nper,pv,start_period,end_period,type)",

    explanation:"Calculates cumulative principal paid during a specified loan period.",

    breakdown:[
        "Rate = Interest Rate",
        "Periods = Loan Duration",
        "Returns Principal Amount"
    ],

    output:"-$50000",

    business:"Used for loan repayment tracking and financial statements.",

    tips:[
        "Works with fixed payment loans.",
        "Useful for amortization reports."
    ],

    errors:[
        "#NUM!",
        "#VALUE!"
    ],

    alternative:"=SUM(PPMT())"
},


// ==========================================================
// FVSCHEDULE
// ==========================================================
{
    title:"FVSCHEDULE",

    keywords:[
        "fvschedule",
        "future value changing rates",
        "variable interest growth",
        "investment growth",
        "multiple interest rates",
        "financial modeling",
        "investment analysis",
        "excel financial function",
        "advanced excel",
        "excel formula"
    ],

    aliases:[
        "variable rate future value",
        "investment growth schedule",
        "compound rate schedule"
    ],

    questions:[
        "How to calculate future value with changing rates?",
        "How to use FVSCHEDULE?"
    ],

    dataset:"investment_schedule",

    category:"Financial",

    formula:"=FVSCHEDULE(10000,B2:B5)",

    syntax:"FVSCHEDULE(principal,schedule)",

    explanation:"Calculates future value using a series of different interest rates.",

    breakdown:[
        "10000 = Initial Investment",
        "B2:B5 = Rate Schedule",
        "Returns Future Value"
    ],

    output:"12500",

    business:"Used for investment models with changing returns.",

    tips:[
        "Useful when rates vary every period.",
        "Alternative to fixed RATE calculations."
    ],

    errors:[
        "#VALUE!"
    ],

    alternative:"=FV()"
},


// ==========================================================
// RRI
// ==========================================================
{
    title:"RRI",

    keywords:[
        "rri",
        "growth rate",
        "investment growth rate",
        "compound growth",
        "annual growth rate",
        "financial analysis",
        "business growth",
        "excel financial function",
        "advanced excel",
        "excel formula"
    ],

    aliases:[
        "compound growth formula",
        "annual growth formula",
        "return growth rate"
    ],

    questions:[
        "How to calculate growth rate?",
        "How to find annual growth?",
        "How to use RRI?"
    ],

    dataset:"investment_growth",

    category:"Financial",

    formula:"=RRI(5,10000,15000)",

    syntax:"RRI(nper,pv,fv)",

    explanation:"Calculates the equivalent growth rate required to reach a future value.",

    breakdown:[
        "5 = Number of Periods",
        "10000 = Present Value",
        "15000 = Future Value"
    ],

    output:"8.45%",

    business:"Used for CAGR and investment growth analysis.",

    tips:[
        "Useful for long-term growth comparison.",
        "Similar to CAGR calculation."
    ],

    errors:[
        "#NUM!",
        "#VALUE!"
    ],

    alternative:"=((FV/PV)^(1/n))-1"
},

// ==========================================================
// DAYS
// ==========================================================
{
    title:"DAYS",

    keywords:[
        "days",
        "date difference",
        "difference between dates",
        "number of days",
        "days between dates",
        "date calculation",
        "aging calculation",
        "sla days",
        "excel date function",
        "advanced excel",
        "mis reporting",
        "dashboard formula",
        "excel formula"
    ],

    aliases:[
        "date gap formula",
        "days calculation",
        "date difference formula"
    ],

    priorityRules:[
        "If query contains working days => NETWORKDAYS.INTL"
    ],

    questions:[
        "How to calculate days between dates?",
        "How to find date difference?",
        "How to use DAYS?"
    ],

    dataset:"leave_records",

    category:"Date & Time",

    formula:"=DAYS(B2,A2)",

    syntax:"DAYS(end_date,start_date)",

    explanation:"Returns the number of days between two dates.",

    breakdown:[
        "B2 = End Date",
        "A2 = Start Date",
        "Returns Total Days"
    ],

    output:"10",

    business:"Used for SLA tracking, ageing reports, and leave calculations.",

    tips:[
        "End date should be greater than start date.",
        "Useful for date ageing analysis."
    ],

    errors:[
        "#VALUE!"
    ],

    alternative:"=B2-A2"
},


// ==========================================================
// DAYS360
// ==========================================================
{
    title:"DAYS360",

    keywords:[
        "days360",
        "360 day calculation",
        "financial date calculation",
        "banking days calculation",
        "year based on 360 days",
        "interest calculation",
        "financial function",
        "advanced excel",
        "excel formula"
    ],

    aliases:[
        "banking date formula",
        "360 calendar formula",
        "financial day count"
    ],

    questions:[
        "How to calculate 360 day difference?",
        "How to use DAYS360?"
    ],

    dataset:"loan_interest",

    category:"Date & Time",

    formula:"=DAYS360(A2,B2)",

    syntax:"DAYS360(start_date,end_date,[method])",

    explanation:"Calculates the difference between two dates based on a 360-day year.",

    breakdown:[
        "A2 = Start Date",
        "B2 = End Date",
        "Uses 360 Day Convention"
    ],

    output:"180",

    business:"Used in banking interest calculations and financial reporting.",

    tips:[
        "Common in financial models.",
        "Supports US and European methods."
    ],

    errors:[
        "#VALUE!"
    ],

    alternative:"=DAYS(B2,A2)"
},


// ==========================================================
// EDATE
// ==========================================================
{
    title:"EDATE",

    keywords:[
        "edate",
        "add months to date",
        "previous month date",
        "future date calculation",
        "date offset",
        "maturity date",
        "renewal date",
        "excel date function",
        "advanced excel",
        "excel formula"
    ],

    aliases:[
        "month addition formula",
        "date after months",
        "month shift formula"
    ],

    questions:[
        "How to add months to date?",
        "How to calculate maturity date?",
        "How to use EDATE?"
    ],

    dataset:"contract_data",

    category:"Date & Time",

    formula:"=EDATE(A2,12)",

    syntax:"EDATE(start_date,months)",

    explanation:"Returns a date after adding or subtracting months.",

    breakdown:[
        "A2 = Starting Date",
        "12 = Months Added",
        "Returns New Date"
    ],

    output:"01-Jan-2027",

    business:"Used for contract expiry, maturity, and renewal tracking.",

    tips:[
        "Negative months return previous dates.",
        "Useful for monthly reporting."
    ],

    errors:[
        "#VALUE!"
    ],

    alternative:"=DATE(YEAR(A2),MONTH(A2)+12,DAY(A2))"
},


// ==========================================================
// EOMONTH
// ==========================================================
{
    title:"EOMONTH",

    keywords:[
        "eomonth",
        "month end date",
        "last date of month",
        "financial month end",
        "closing date",
        "month closing",
        "reporting date",
        "excel date function",
        "advanced excel",
        "excel formula"
    ],

    aliases:[
        "month end formula",
        "closing date formula",
        "last day formula"
    ],

    questions:[
        "How to find month end date?",
        "How to calculate closing date?",
        "How to use EOMONTH?"
    ],

    dataset:"monthly_reporting",

    category:"Date & Time",

    formula:"=EOMONTH(A2,0)",

    syntax:"EOMONTH(start_date,months)",

    explanation:"Returns the last day of a month before or after a specified date.",

    breakdown:[
        "A2 = Reference Date",
        "0 = Same Month",
        "Returns Month End"
    ],

    output:"31-Jul-2026",

    business:"Used for MIS closing, month-end reports, and financial statements.",

    tips:[
        "Very useful in monthly automation.",
        "Works with fiscal reporting."
    ],

    errors:[
        "#VALUE!"
    ],

    alternative:"=DATE(YEAR(A2),MONTH(A2)+1,0)"
},


// ==========================================================
// NETWORKDAYS.INTL
// ==========================================================
{
    title:"NETWORKDAYS.INTL",

    keywords:[
        "networkdays.intl",
        "working days",
        "business days",
        "exclude weekends",
        "exclude holidays",
        "sla calculation",
        "leave calculation",
        "working day formula",
        "excel date function",
        "advanced excel",
        "mis reporting",
        "excel formula"
    ],

    aliases:[
        "working days formula",
        "business day calculator",
        "holiday excluding days"
    ],

    priorityRules:[
        "If query contains simple working days => NETWORKDAYS"
    ],

    questions:[
        "How to calculate working days?",
        "How to exclude weekends?",
        "How to use NETWORKDAYS.INTL?"
    ],

    dataset:"employee_leave",

    category:"Date & Time",

    formula:"=NETWORKDAYS.INTL(A2,B2,1,D2:D10)",

    syntax:"NETWORKDAYS.INTL(start_date,end_date,[weekend],[holidays])",

    explanation:"Calculates working days between dates with custom weekends and holidays.",

    breakdown:[
        "A2 = Start Date",
        "B2 = End Date",
        "D2:D10 = Holiday List"
    ],

    output:"8",

    business:"Used for SLA tracking, leave management, and operational reports.",

    tips:[
        "Supports custom weekend patterns.",
        "Ideal for HR and compliance reports."
    ],

    errors:[
        "#VALUE!"
    ],

    alternative:"=NETWORKDAYS(A2,B2)"
},

// ==========================================================
// LET
// ==========================================================
{
    title:"LET",

    keywords:[
        "let",
        "store calculation",
        "formula variable",
        "simplify formula",
        "replace repeated calculation",
        "advanced formula",
        "excel 365",
        "formula optimization",
        "dynamic calculation",
        "excel formula"
    ],

    aliases:[
        "variable formula",
        "named calculation formula",
        "optimized formula"
    ],

    questions:[
        "How to create variables in formula?",
        "How to simplify long formulas?",
        "How to use LET?"
    ],

    dataset:"sales_analysis",

    category:"Advanced Array",

    formula:"=LET(total,B2*C2,total*10%)",

    syntax:"LET(name1,name_value,calculation)",

    explanation:"Allows you to create variables inside a formula to improve readability and performance.",

    breakdown:[
        "total = Variable Name",
        "B2*C2 = Stored Calculation",
        "Returns Final Result"
    ],

    output:"5000",

    business:"Used in complex MIS formulas and dashboard calculations.",

    tips:[
        "Reduces repeated calculations.",
        "Available in Excel 365."
    ],

    errors:[
        "#NAME!",
        "#VALUE!"
    ],

    alternative:"=B2*C2*10%"
},


// ==========================================================
// LAMBDA
// ==========================================================
{
    title:"LAMBDA",

    keywords:[
        "lambda",
        "custom function",
        "create own function",
        "reusable formula",
        "user defined formula",
        "advanced excel function",
        "excel 365",
        "automation",
        "excel formula"
    ],

    aliases:[
        "custom excel function",
        "formula function creator",
        "user formula"
    ],

    questions:[
        "How to create custom Excel function?",
        "How to make reusable formula?",
        "How to use LAMBDA?"
    ],

    dataset:"calculation_model",

    category:"Advanced Array",

    formula:"=LAMBDA(x,x*10)(5)",

    syntax:"LAMBDA(parameter,calculation)",

    explanation:"Creates custom reusable Excel functions without VBA.",

    breakdown:[
        "x = Input Parameter",
        "x*10 = Calculation",
        "Returns Custom Result"
    ],

    output:"50",

    business:"Used for automation and advanced Excel solutions.",

    tips:[
        "Can be saved in Name Manager.",
        "Works like a custom function."
    ],

    errors:[
        "#NAME!",
        "#CALC!"
    ],

    alternative:"VBA Function"
},


// ==========================================================
// MAP
// ==========================================================
{
    title:"MAP",

    keywords:[
        "map",
        "apply formula each value",
        "array calculation",
        "dynamic array",
        "transform data",
        "excel 365",
        "advanced excel",
        "automation",
        "excel formula"
    ],

    aliases:[
        "array mapping formula",
        "apply function to range",
        "dynamic calculation"
    ],

    questions:[
        "How to apply formula to every value?",
        "How to transform array?",
        "How to use MAP?"
    ],

    dataset:"employee_scores",

    category:"Advanced Array",

    formula:"=MAP(B2:B10,LAMBDA(x,x*10%))",

    syntax:"MAP(array,LAMBDA)",

    explanation:"Applies a custom calculation to each value in an array.",

    breakdown:[
        "B2:B10 = Input Array",
        "LAMBDA = Calculation",
        "Returns Transformed Array"
    ],

    output:"90,85,95",

    business:"Used for large-scale automated calculations.",

    tips:[
        "Works with dynamic arrays.",
        "Requires Excel 365."
    ],

    errors:[
        "#CALC!",
        "#VALUE!"
    ],

    alternative:"=B2:B10*10%"
},


// ==========================================================
// REDUCE
// ==========================================================
{
    title:"REDUCE",

    keywords:[
        "reduce",
        "combine values",
        "aggregate array",
        "accumulate result",
        "array calculation",
        "excel 365",
        "advanced excel",
        "excel formula"
    ],

    aliases:[
        "array accumulator",
        "combine array formula",
        "running total function"
    ],

    questions:[
        "How to combine array values?",
        "How to accumulate results?",
        "How to use REDUCE?"
    ],

    dataset:"transaction_data",

    category:"Advanced Array",

    formula:"=REDUCE(0,B2:B10,LAMBDA(a,b,a+b))",

    syntax:"REDUCE(initial,array,LAMBDA)",

    explanation:"Reduces an array into a single accumulated value.",

    breakdown:[
        "0 = Starting Value",
        "B2:B10 = Array",
        "Adds Values Sequentially"
    ],

    output:"50000",

    business:"Used for complex aggregation and automation models.",

    tips:[
        "Similar to looping without VBA.",
        "Available in Excel 365."
    ],

    errors:[
        "#CALC!",
        "#VALUE!"
    ],

    alternative:"=SUM(B2:B10)"
},


// ==========================================================
// SCAN
// ==========================================================
{
    title:"SCAN",

    keywords:[
        "scan",
        "running calculation",
        "running total",
        "cumulative result",
        "array tracking",
        "dynamic array",
        "excel 365",
        "advanced excel",
        "excel formula"
    ],

    aliases:[
        "running total formula",
        "cumulative calculation",
        "array accumulation"
    ],

    questions:[
        "How to create running total?",
        "How to track cumulative values?",
        "How to use SCAN?"
    ],

    dataset:"monthly_sales",

    category:"Advanced Array",

    formula:"=SCAN(0,B2:B10,LAMBDA(a,b,a+b))",

    syntax:"SCAN(initial,array,LAMBDA)",

    explanation:"Returns intermediate results while performing calculations across an array.",

    breakdown:[
        "0 = Initial Value",
        "B2:B10 = Values",
        "Returns Running Results"
    ],

    output:"1000,2500,4000",

    business:"Used for cumulative reports and trend analysis.",

    tips:[
        "Unlike REDUCE, returns every intermediate value.",
        "Useful for dashboards."
    ],

    errors:[
        "#CALC!",
        "#VALUE!"
    ],

    alternative:"=SUM($B$2:B2)"
},

// ==========================================================
// FORMULATEXT
// ==========================================================
{
    title:"FORMULATEXT",

    keywords:[
        "formulatext",
        "show formula text",
        "display formula",
        "extract formula",
        "formula documentation",
        "formula audit",
        "excel auditing",
        "advanced excel",
        "excel formula"
    ],

    aliases:[
        "formula viewer",
        "formula extraction",
        "formula display"
    ],

    questions:[
        "How to show formula as text?",
        "How to extract formula?",
        "How to use FORMULATEXT?"
    ],

    dataset:"formula_audit",

    category:"Information",

    formula:"=FORMULATEXT(B2)",

    syntax:"FORMULATEXT(reference)",

    explanation:"Returns the formula contained in a referenced cell as text.",

    breakdown:[
        "B2 = Formula Cell",
        "Reads Formula",
        "Returns Formula Text"
    ],

    output:"=SUM(A1:A10)",

    business:"Used for formula auditing and documentation tools.",

    tips:[
        "Does not calculate the formula.",
        "Useful for Excel AI assistants."
    ],

    errors:[
        "#N/A",
        "#VALUE!"
    ],

    alternative:"Copy Formula Manually"
},


// ==========================================================
// ISFORMULA
// ==========================================================
{
    title:"ISFORMULA",

    keywords:[
        "isformula",
        "check formula",
        "identify formula cells",
        "formula validation",
        "audit formula",
        "excel checking",
        "information function",
        "advanced excel",
        "excel formula"
    ],

    aliases:[
        "formula checker",
        "formula detection",
        "formula validation"
    ],

    questions:[
        "How to check if cell contains formula?",
        "How to identify formulas?",
        "How to use ISFORMULA?"
    ],

    dataset:"worksheet_audit",

    category:"Information",

    formula:"=ISFORMULA(A2)",

    syntax:"ISFORMULA(reference)",

    explanation:"Checks whether a cell contains a formula.",

    breakdown:[
        "A2 = Target Cell",
        "Checks Formula Presence",
        "Returns TRUE/FALSE"
    ],

    output:"TRUE",

    business:"Used for audit checks and spreadsheet quality control.",

    tips:[
        "Works with conditional formatting.",
        "Useful for protecting formulas."
    ],

    errors:[
        "#VALUE!"
    ],

    alternative:"=FORMULATEXT(A2)"
},


// ==========================================================
// CELL
// ==========================================================
{
    title:"CELL",

    keywords:[
        "cell",
        "cell information",
        "get cell details",
        "cell address",
        "format information",
        "excel information function",
        "advanced excel",
        "excel formula"
    ],

    aliases:[
        "cell metadata formula",
        "cell property formula",
        "cell details"
    ],

    questions:[
        "How to get cell information?",
        "How to find cell address?",
        "How to use CELL?"
    ],

    dataset:"worksheet_information",

    category:"Information",

    formula:'=CELL("address",A1)',

    syntax:"CELL(info_type,[reference])",

    explanation:"Returns information about the formatting, location, or contents of a cell.",

    breakdown:[
        "address = Information Type",
        "A1 = Reference Cell",
        "Returns Cell Details"
    ],

    output:"$A$1",

    business:"Used for dynamic reporting and workbook analysis.",

    tips:[
        "Supports multiple information types.",
        "Useful in dashboard automation."
    ],

    errors:[
        "#VALUE!"
    ],

    alternative:"=ADDRESS()"
},


// ==========================================================
// INFO
// ==========================================================
{
    title:"INFO",

    keywords:[
        "info",
        "excel system information",
        "environment details",
        "workbook information",
        "system check",
        "information function",
        "advanced excel",
        "excel formula"
    ],

    aliases:[
        "system information formula",
        "excel environment details"
    ],

    questions:[
        "How to get Excel system information?",
        "How to use INFO?"
    ],

    dataset:"system_information",

    category:"Information",

    formula:'=INFO("system")',

    syntax:"INFO(type_text)",

    explanation:"Returns information about the operating environment of Excel.",

    breakdown:[
        "system = Information Type",
        "Reads Excel Environment",
        "Returns System Details"
    ],

    output:"pcdos",

    business:"Used for technical checks and workbook diagnostics.",

    tips:[
        "Limited use in modern Excel.",
        "Useful for compatibility checks."
    ],

    errors:[
        "#VALUE!"
    ],

    alternative:"=CELL()"
},


// ==========================================================
// STOCKHISTORY
// ==========================================================
{
    title:"STOCKHISTORY",

    keywords:[
        "stockhistory",
        "stock price history",
        "share price",
        "market data",
        "historical stock price",
        "investment analysis",
        "financial data",
        "excel finance function",
        "advanced excel",
        "excel formula"
    ],

    aliases:[
        "stock data formula",
        "market history formula",
        "share price tracker"
    ],

    questions:[
        "How to get historical stock prices?",
        "How to track share prices?",
        "How to use STOCKHISTORY?"
    ],

    dataset:"market_data",

    category:"Financial",

    formula:'=STOCKHISTORY("MSFT",DATE(2025,1,1),DATE(2025,12,31))',

    syntax:"STOCKHISTORY(stock,start_date,end_date,[interval],[headers],[properties])",

    explanation:"Returns historical stock price data from supported financial markets.",

    breakdown:[
        "MSFT = Stock Symbol",
        "Start Date = Beginning Date",
        "End Date = Closing Date"
    ],

    output:"Historical Price Table",

    business:"Used for investment dashboards and market analysis.",

    tips:[
        "Requires Excel 365 connected data features.",
        "Useful for finance dashboards."
    ],

    errors:[
        "#N/A",
        "#VALUE!"
    ],

    alternative:"Power Query Market Data Import"
},

// ==========================================================
// SLA_BREACH
// ==========================================================
{
    title:"SLA_BREACH",

    keywords:[
        "sla breach",
        "sla status",
        "service level agreement",
        "breach calculation",
        "deadline check",
        "operation monitoring",
        "compliance reporting",
        "mis reporting",
        "dashboard formula",
        "excel formula"
    ],

    aliases:[
        "sla tracker formula",
        "deadline breach formula",
        "service level check"
    ],

    questions:[
        "How to calculate SLA breach?",
        "How to identify delayed cases?",
        "How to track SLA status?"
    ],

    dataset:"sla_tracker",

    category:"Banking MIS",

    formula:'=IF(D2>C2,"BREACHED","WITHIN SLA")',

    syntax:"IF(condition,true_result,false_result)",

    explanation:"Checks whether actual completion date exceeds SLA deadline.",

    breakdown:[
        "D2 = Completion Date",
        "C2 = SLA Due Date",
        "Compares Delay Status"
    ],

    output:"BREACHED",

    business:"Used for operational SLA monitoring and compliance dashboards.",

    tips:[
        "Combine with conditional formatting.",
        "Useful for daily MIS reports."
    ],

    errors:[
        "#VALUE!"
    ],

    alternative:'=IF(D2-C2>0,"BREACHED","OK")'
},


// ==========================================================
// AGEING_BUCKET
// ==========================================================
{
    title:"AGEING_BUCKET",

    keywords:[
        "ageing",
        "aging report",
        "pending days",
        "bucket analysis",
        "outstanding ageing",
        "open items",
        "receivable ageing",
        "mis reporting",
        "dashboard formula",
        "excel formula"
    ],

    aliases:[
        "age bucket formula",
        "pending days classification",
        "outstanding analysis"
    ],

    questions:[
        "How to create ageing buckets?",
        "How to classify pending days?",
        "How to make ageing report?"
    ],

    dataset:"pending_cases",

    category:"Banking MIS",

    formula:'=IF(TODAY()-A2<=7,"0-7 Days",IF(TODAY()-A2<=30,"8-30 Days","30+ Days"))',

    syntax:"IF(logical_test,value_if_true,value_if_false)",

    explanation:"Categorizes records based on pending days.",

    breakdown:[
        "A2 = Open Date",
        "TODAY() = Current Date",
        "Returns Ageing Bucket"
    ],

    output:"8-30 Days",

    business:"Used for pending cases, audit observations, and operational ageing.",

    tips:[
        "Create custom buckets as per business requirement.",
        "Useful in dashboards."
    ],

    errors:[
        "#VALUE!"
    ],

    alternative:"=IFS()"
},


// ==========================================================
// TAT_CALCULATION
// ==========================================================
{
    title:"TAT_CALCULATION",

    keywords:[
        "tat",
        "turnaround time",
        "processing time",
        "completion time",
        "case duration",
        "operations report",
        "mis reporting",
        "dashboard formula",
        "excel formula"
    ],

    aliases:[
        "time taken formula",
        "processing days formula",
        "duration calculation"
    ],

    questions:[
        "How to calculate turnaround time?",
        "How to calculate processing days?",
        "How to track TAT?"
    ],

    dataset:"operation_cases",

    category:"Banking MIS",

    formula:"=B2-A2",

    syntax:"End Date - Start Date",

    explanation:"Calculates the total time taken to complete a task.",

    breakdown:[
        "A2 = Start Date",
        "B2 = Completion Date",
        "Returns Duration"
    ],

    output:"5 Days",

    business:"Used for operational efficiency and SLA reporting.",

    tips:[
        "Use NETWORKDAYS.INTL for working days.",
        "Exclude holidays when required."
    ],

    errors:[
        "#VALUE!"
    ],

    alternative:"=NETWORKDAYS(A2,B2)"
},


// ==========================================================
// COMPLIANCE_STATUS
// ==========================================================
{
    title:"COMPLIANCE_STATUS",

    keywords:[
        "compliance",
        "compliance check",
        "policy validation",
        "control testing",
        "audit status",
        "rbi compliance",
        "risk monitoring",
        "mis reporting",
        "excel formula"
    ],

    aliases:[
        "control check formula",
        "audit validation",
        "policy status"
    ],

    questions:[
        "How to create compliance status?",
        "How to check control compliance?",
        "How to automate audit status?"
    ],

    dataset:"compliance_tracker",

    category:"Banking MIS",

    formula:'=IF(AND(B2="Yes",C2="Yes"),"Compliant","Exception")',

    syntax:"IF(AND(condition1,condition2),result1,result2)",

    explanation:"Checks multiple compliance conditions and returns status.",

    breakdown:[
        "B2 = Control Completed",
        "C2 = Evidence Available",
        "Returns Compliance Status"
    ],

    output:"Compliant",

    business:"Used for audit, ICOFR, SOX, and compliance monitoring.",

    tips:[
        "Add more controls using AND.",
        "Useful for control dashboards."
    ],

    errors:[
        "#VALUE!"
    ],

    alternative:'=IFS()'
},


// ==========================================================
// EXCEPTION_FLAG
// ==========================================================
{
    title:"EXCEPTION_FLAG",

    keywords:[
        "exception",
        "exception report",
        "error flag",
        "identify issue",
        "risk indicator",
        "control failure",
        "data validation",
        "mis reporting",
        "dashboard formula",
        "excel formula"
    ],

    aliases:[
        "issue flag formula",
        "error identification",
        "risk flag"
    ],

    questions:[
        "How to highlight exceptions?",
        "How to create error flags?",
        "How to identify issues automatically?"
    ],

    dataset:"control_testing",

    category:"Banking MIS",

    formula:'=IF(B2<>"OK","CHECK","CLEAR")',

    syntax:"IF(condition,true_result,false_result)",

    explanation:"Creates an automated flag for exception identification.",

    breakdown:[
        "B2 = Control Result",
        "Checks Exception Condition",
        "Returns Flag"
    ],

    output:"CHECK",

    business:"Used for audit exceptions, operational issues, and risk reporting.",

    tips:[
        "Combine with conditional formatting.",
        "Useful for automated MIS."
    ],

    errors:[
        "#VALUE!"
    ],

    alternative:'=IFERROR()'
},

// ==========================================================
// ATTRITION_RATE
// ==========================================================
{
    title:"ATTRITION_RATE",

    keywords:[
        "attrition",
        "employee attrition",
        "employee turnover",
        "staff leaving",
        "resignation analysis",
        "hr analytics",
        "workforce analysis",
        "mis reporting",
        "dashboard formula",
        "excel formula"
    ],

    aliases:[
        "employee turnover formula",
        "exit percentage formula",
        "hr attrition calculation"
    ],

    questions:[
        "How to calculate employee attrition?",
        "How to find turnover percentage?",
        "How to create attrition report?"
    ],

    dataset:"employee_hr_data",

    category:"HR Analytics",

    formula:"=B2/C2*100",

    syntax:"Employees Left / Average Employees * 100",

    explanation:"Calculates employee attrition percentage.",

    breakdown:[
        "B2 = Number of Employees Left",
        "C2 = Average Employee Strength",
        "Returns Attrition %"
    ],

    output:"5%",

    business:"Used for HR dashboards and workforce planning.",

    tips:[
        "Calculate monthly or yearly attrition.",
        "Useful for management reporting."
    ],

    errors:[
        "#DIV/0!"
    ],

    alternative:"=COUNTIF()"
},


// ==========================================================
// ATTENDANCE_PERCENTAGE
// ==========================================================
{
    title:"ATTENDANCE_PERCENTAGE",

    keywords:[
        "attendance percentage",
        "employee attendance",
        "present percentage",
        "working days",
        "attendance report",
        "hr dashboard",
        "employee analytics",
        "excel formula"
    ],

    aliases:[
        "attendance rate formula",
        "presence percentage",
        "attendance calculation"
    ],

    questions:[
        "How to calculate attendance percentage?",
        "How to create attendance report?",
        "How to calculate present days?"
    ],

    dataset:"attendance_register",

    category:"HR Analytics",

    formula:"=B2/C2*100",

    syntax:"Present Days / Total Working Days * 100",

    explanation:"Calculates employee attendance percentage.",

    breakdown:[
        "B2 = Present Days",
        "C2 = Total Working Days",
        "Returns Attendance %"
    ],

    output:"96%",

    business:"Used for employee attendance monitoring.",

    tips:[
        "Exclude approved leave if required.",
        "Useful for monthly HR MIS."
    ],

    errors:[
        "#DIV/0!"
    ],

    alternative:"=COUNTIF()"
},


// ==========================================================
// LEAVE_BALANCE
// ==========================================================
{
    title:"LEAVE_BALANCE",

    keywords:[
        "leave balance",
        "remaining leave",
        "available leaves",
        "employee leave",
        "leave calculation",
        "hr reporting",
        "attendance management",
        "excel formula"
    ],

    aliases:[
        "leave remaining formula",
        "leave tracker",
        "leave availability"
    ],

    questions:[
        "How to calculate leave balance?",
        "How to track remaining leaves?",
        "How to create leave report?"
    ],

    dataset:"employee_leave",

    category:"HR Analytics",

    formula:"=B2-C2",

    syntax:"Allocated Leave - Used Leave",

    explanation:"Calculates remaining employee leave balance.",

    breakdown:[
        "B2 = Total Allocated Leave",
        "C2 = Used Leave",
        "Returns Balance Leave"
    ],

    output:"8 Days",

    business:"Used for employee leave management and HR dashboards.",

    tips:[
        "Can be extended for leave encashment.",
        "Works with yearly leave cycles."
    ],

    errors:[
        "#VALUE!"
    ],

    alternative:"=SUMIFS()"
},


// ==========================================================
// EMPLOYEE_TENURE
// ==========================================================
{
    title:"EMPLOYEE_TENURE",

    keywords:[
        "employee tenure",
        "years of service",
        "joining date calculation",
        "employee experience",
        "service period",
        "hr analytics",
        "employee profile",
        "excel formula"
    ],

    aliases:[
        "experience calculation",
        "service year formula",
        "employment duration"
    ],

    questions:[
        "How to calculate employee tenure?",
        "How to find years of service?",
        "How to calculate experience?"
    ],

    dataset:"employee_master",

    category:"HR Analytics",

    formula:'=DATEDIF(A2,TODAY(),"Y")',

    syntax:"DATEDIF(start_date,end_date,unit)",

    explanation:"Calculates completed years of employee service.",

    breakdown:[
        "A2 = Joining Date",
        "TODAY() = Current Date",
        "Y = Years Difference"
    ],

    output:"8 Years",

    business:"Used for workforce segmentation and experience analysis.",

    tips:[
        "Can calculate months and days also.",
        "Useful for HR reports."
    ],

    errors:[
        "#NUM!",
        "#VALUE!"
    ],

    alternative:"=YEARFRAC()"
},


// ==========================================================
// PRODUCTIVITY_SCORE
// ==========================================================
{
    title:"PRODUCTIVITY_SCORE",

    keywords:[
        "productivity score",
        "employee performance",
        "work efficiency",
        "task completion",
        "employee ranking",
        "performance dashboard",
        "hr analytics",
        "excel formula"
    ],

    aliases:[
        "employee efficiency formula",
        "performance score",
        "work output score"
    ],

    questions:[
        "How to calculate productivity score?",
        "How to measure employee efficiency?",
        "How to create performance dashboard?"
    ],

    dataset:"employee_productivity",

    category:"HR Analytics",

    formula:"=B2/C2*100",

    syntax:"Completed Tasks / Assigned Tasks * 100",

    explanation:"Calculates employee productivity percentage.",

    breakdown:[
        "B2 = Completed Tasks",
        "C2 = Assigned Tasks",
        "Returns Productivity Score"
    ],

    output:"92%",

    business:"Used for employee performance analytics and MIS dashboards.",

    tips:[
        "Can include quality score weighting.",
        "Useful for team comparison."
    ],

    errors:[
        "#DIV/0!"
    ],

    alternative:"=AVERAGEIFS()"
},

// ==========================================================
// GROSS_MARGIN_PERCENTAGE
// ==========================================================
{
    title:"GROSS_MARGIN_PERCENTAGE",

    keywords:[
        "gross margin",
        "gross profit percentage",
        "profit margin",
        "revenue analysis",
        "finance dashboard",
        "profitability analysis",
        "financial reporting",
        "mis reporting",
        "excel formula"
    ],

    aliases:[
        "gross profit formula",
        "margin percentage",
        "business margin"
    ],

    questions:[
        "How to calculate gross margin?",
        "How to find gross profit percentage?",
        "How to create profitability report?"
    ],

    dataset:"finance_statement",

    category:"Finance & Risk",

    formula:"=(B2-C2)/B2*100",

    syntax:"(Revenue - Cost) / Revenue * 100",

    explanation:"Calculates the percentage of revenue remaining after deducting direct costs.",

    breakdown:[
        "B2 = Revenue",
        "C2 = Cost of Goods Sold",
        "Returns Gross Margin %"
    ],

    output:"40%",

    business:"Used for profitability analysis and financial MIS reporting.",

    tips:[
        "Higher margin indicates better profitability.",
        "Useful for product-level analysis."
    ],

    errors:[
        "#DIV/0!"
    ],

    alternative:"=((Revenue-Cost)/Revenue)"
},


// ==========================================================
// NET_PROFIT_MARGIN
// ==========================================================
{
    title:"NET_PROFIT_MARGIN",

    keywords:[
        "net profit margin",
        "profit percentage",
        "bottom line",
        "financial ratio",
        "profitability",
        "finance analysis",
        "management reporting",
        "excel formula"
    ],

    aliases:[
        "net income percentage",
        "profit ratio formula",
        "earnings margin"
    ],

    questions:[
        "How to calculate net profit margin?",
        "How to find profitability ratio?",
        "How to create finance dashboard?"
    ],

    dataset:"profit_loss_statement",

    category:"Finance & Risk",

    formula:"=B2/C2*100",

    syntax:"Net Profit / Revenue * 100",

    explanation:"Calculates the percentage of revenue converted into net profit.",

    breakdown:[
        "B2 = Net Profit",
        "C2 = Revenue",
        "Returns Net Margin %"
    ],

    output:"15%",

    business:"Used for financial performance reporting.",

    tips:[
        "Compare across periods for trend analysis.",
        "Useful for management review."
    ],

    errors:[
        "#DIV/0!"
    ],

    alternative:"=Profit/Revenue"
},


// ==========================================================
// VARIANCE_ANALYSIS
// ==========================================================
{
    title:"VARIANCE_ANALYSIS",

    keywords:[
        "variance",
        "actual vs budget",
        "difference analysis",
        "budget variance",
        "financial comparison",
        "forecast analysis",
        "management reporting",
        "excel formula"
    ],

    aliases:[
        "difference formula",
        "budget gap formula",
        "actual variance"
    ],

    questions:[
        "How to calculate variance?",
        "How to compare actual and budget?",
        "How to create variance report?"
    ],

    dataset:"budget_report",

    category:"Finance & Risk",

    formula:"=B2-C2",

    syntax:"Actual - Budget",

    explanation:"Calculates the difference between actual and planned values.",

    breakdown:[
        "B2 = Actual Amount",
        "C2 = Budget Amount",
        "Returns Variance"
    ],

    output:"5000",

    business:"Used for budget monitoring and financial control reports.",

    tips:[
        "Add percentage variance for better analysis.",
        "Positive/negative meaning depends on metric."
    ],

    errors:[
        "#VALUE!"
    ],

    alternative:"=(Actual-Budget)/Budget"
},


// ==========================================================
// BUDGET_UTILIZATION
// ==========================================================
{
    title:"BUDGET_UTILIZATION",

    keywords:[
        "budget utilization",
        "budget consumed",
        "expense utilization",
        "budget tracking",
        "financial control",
        "spending percentage",
        "finance dashboard",
        "excel formula"
    ],

    aliases:[
        "budget usage formula",
        "expense utilization rate",
        "budget consumption"
    ],

    questions:[
        "How to calculate budget utilization?",
        "How much budget is used?",
        "How to track expenses?"
    ],

    dataset:"expense_tracking",

    category:"Finance & Risk",

    formula:"=B2/C2*100",

    syntax:"Used Budget / Total Budget * 100",

    explanation:"Calculates percentage of budget consumed.",

    breakdown:[
        "B2 = Used Amount",
        "C2 = Total Budget",
        "Returns Utilization %"
    ],

    output:"75%",

    business:"Used for expense monitoring and budget control.",

    tips:[
        "Use conditional formatting for overspending.",
        "Useful in finance dashboards."
    ],

    errors:[
        "#DIV/0!"
    ],

    alternative:"=SUMIFS()"
},


// ==========================================================
// RISK_SCORE
// ==========================================================
{
    title:"RISK_SCORE",

    keywords:[
        "risk score",
        "risk rating",
        "risk calculation",
        "compliance risk",
        "control risk",
        "risk dashboard",
        "audit analysis",
        "excel formula"
    ],

    aliases:[
        "risk rating formula",
        "risk indicator",
        "control risk score"
    ],

    questions:[
        "How to calculate risk score?",
        "How to create risk rating?",
        "How to automate risk classification?"
    ],

    dataset:"risk_register",

    category:"Finance & Risk",

    formula:'=B2*C2',

    syntax:"Risk Probability * Risk Impact",

    explanation:"Calculates risk score based on probability and impact.",

    breakdown:[
        "B2 = Probability Score",
        "C2 = Impact Score",
        "Returns Risk Score"
    ],

    output:"15",

    business:"Used in risk management, audit, and compliance dashboards.",

    tips:[
        "Define risk rating thresholds separately.",
        "Useful for RCSA and control testing."
    ],

    errors:[
        "#VALUE!"
    ],

    alternative:"=IFS()"
},

// ==========================================================
// DUPLICATE_DETECTION
// ==========================================================
{
    title:"DUPLICATE_DETECTION",

    keywords:[
        "duplicate",
        "find duplicate records",
        "duplicate check",
        "repeated values",
        "data validation",
        "data cleaning",
        "data quality",
        "excel formula",
        "data analyst"
    ],

    aliases:[
        "duplicate finder",
        "duplicate flag",
        "repeat value check"
    ],

    questions:[
        "How to find duplicate records?",
        "How to identify repeated values?",
        "How to create duplicate flag?"
    ],

    dataset:"customer_data",

    category:"Data Analytics",

    formula:'=IF(COUNTIF(A:A,A2)>1,"Duplicate","Unique")',

    syntax:"IF(COUNTIF(range,value)>1,result1,result2)",

    explanation:"Identifies duplicate values in a dataset.",

    breakdown:[
        "A:A = Search Column",
        "A2 = Current Value",
        "COUNTIF Counts Occurrences"
    ],

    output:"Duplicate",

    business:"Used for customer master cleanup and data validation.",

    tips:[
        "Useful before reporting.",
        "Can be combined with conditional formatting."
    ],

    errors:[
        "#VALUE!"
    ],

    alternative:"Remove Duplicates Feature"
},


// ==========================================================
// DATA_QUALITY_SCORE
// ==========================================================
{
    title:"DATA_QUALITY_SCORE",

    keywords:[
        "data quality score",
        "data accuracy",
        "data completeness",
        "quality dashboard",
        "data validation",
        "analytics reporting",
        "data analyst",
        "excel formula"
    ],

    aliases:[
        "data health score",
        "quality percentage",
        "dataset quality"
    ],

    questions:[
        "How to calculate data quality score?",
        "How to measure data accuracy?",
        "How to create data health dashboard?"
    ],

    dataset:"data_quality_check",

    category:"Data Analytics",

    formula:"=B2/C2*100",

    syntax:"Valid Records / Total Records * 100",

    explanation:"Calculates the percentage of valid records in a dataset.",

    breakdown:[
        "B2 = Valid Records",
        "C2 = Total Records",
        "Returns Quality Percentage"
    ],

    output:"98%",

    business:"Used for data governance and reporting quality checks.",

    tips:[
        "Add multiple quality parameters.",
        "Useful in ETL monitoring."
    ],

    errors:[
        "#DIV/0!"
    ],

    alternative:"=AVERAGE()"
},


// ==========================================================
// MISSING_VALUE_PERCENTAGE
// ==========================================================
{
    title:"MISSING_VALUE_PERCENTAGE",

    keywords:[
        "missing value",
        "blank percentage",
        "null check",
        "empty data",
        "data completeness",
        "data cleaning",
        "data analyst",
        "excel formula"
    ],

    aliases:[
        "blank percentage formula",
        "missing data analysis",
        "null percentage"
    ],

    questions:[
        "How to calculate missing values?",
        "How to find blank percentage?",
        "How to check data completeness?"
    ],

    dataset:"raw_customer_data",

    category:"Data Analytics",

    formula:'=COUNTBLANK(A2:A100)/COUNTA(A2:A100)*100',

    syntax:"Blank Cells / Total Cells * 100",

    explanation:"Calculates percentage of missing values in a dataset.",

    breakdown:[
        "COUNTBLANK = Missing Records",
        "COUNTA = Total Records",
        "Returns Missing Percentage"
    ],

    output:"2%",

    business:"Used for data cleaning before analysis.",

    tips:[
        "Important before creating dashboards.",
        "Track data completeness regularly."
    ],

    errors:[
        "#DIV/0!"
    ],

    alternative:"Power Query Column Quality"
},


// ==========================================================
// KPI_ACHIEVEMENT_PERCENTAGE
// ==========================================================
{
    title:"KPI_ACHIEVEMENT_PERCENTAGE",

    keywords:[
        "kpi achievement",
        "target achieved",
        "performance percentage",
        "goal completion",
        "business kpi",
        "dashboard metric",
        "management reporting",
        "excel formula"
    ],

    aliases:[
        "target completion formula",
        "achievement rate",
        "performance score"
    ],

    questions:[
        "How to calculate KPI achievement?",
        "How to measure target completion?",
        "How to create KPI dashboard?"
    ],

    dataset:"performance_tracker",

    category:"Data Analytics",

    formula:"=B2/C2*100",

    syntax:"Actual Result / Target * 100",

    explanation:"Calculates percentage of KPI achievement against target.",

    breakdown:[
        "B2 = Actual Achievement",
        "C2 = Target Value",
        "Returns Achievement %"
    ],

    output:"92%",

    business:"Used for management dashboards and performance tracking.",

    tips:[
        "Add status using IF or IFS.",
        "Useful for scorecards."
    ],

    errors:[
        "#DIV/0!"
    ],

    alternative:"=IF(B2>=C2,\"Achieved\",\"Pending\")"
},


// ==========================================================
// RANKING_ANALYSIS
// ==========================================================
{
    title:"RANKING_ANALYSIS",

    keywords:[
        "ranking",
        "rank employees",
        "top performer",
        "performance ranking",
        "leaderboard",
        "score ranking",
        "analytics dashboard",
        "excel formula"
    ],

    aliases:[
        "rank formula",
        "top performer formula",
        "position calculation"
    ],

    questions:[
        "How to rank employees?",
        "How to create leaderboard?",
        "How to find top performers?"
    ],

    dataset:"employee_performance",

    category:"Data Analytics",

    formula:"=RANK.EQ(B2,$B$2:$B$100,0)",

    syntax:"RANK.EQ(number,ref,[order])",

    explanation:"Assigns a ranking position based on values.",

    breakdown:[
        "B2 = Performance Score",
        "$B$2:$B$100 = Comparison Range",
        "0 = Highest Value First"
    ],

    output:"Rank 1",

    business:"Used for employee ranking, sales leaderboard, and KPI reports.",

    tips:[
        "Use RANK.AVG when duplicates exist.",
        "Combine with SORT for dashboards."
    ],

    errors:[
        "#N/A",
        "#VALUE!"
    ],

    alternative:"=SORT()"
},

// ==========================================================
// FORMULA_RECOMMENDATION
// ==========================================================
{
    title:"FORMULA_RECOMMENDATION",

    keywords:[
        "suggest formula",
        "formula recommendation",
        "which formula to use",
        "excel ai",
        "formula finder",
        "formula assistant",
        "excel helper",
        "automation"
    ],

    aliases:[
        "formula suggestion",
        "ai formula generator",
        "formula advisor"
    ],

    questions:[
        "Which formula should I use?",
        "How to get formula suggestion?",
        "How does Excel AI recommend formulas?"
    ],

    dataset:"formula_library",

    category:"AI Assistant",

    formula:'=IF(A2="Total","SUM",IF(A2="Lookup","XLOOKUP","Check Requirement"))',

    syntax:"IF(condition,result1,result2)",

    explanation:"Suggests a suitable Excel function based on user requirement.",

    breakdown:[
        "A2 = User Requirement",
        "Checks Keyword",
        "Returns Recommended Formula"
    ],

    output:"XLOOKUP",

    business:"Used in AI-powered Excel formula recommendation tools.",

    tips:[
        "Combine with keyword database.",
        "Use JavaScript search engine for better results."
    ],

    errors:[
        "#N/A"
    ],

    alternative:"AI Search Engine"
},


// ==========================================================
// ERROR_DETECTION
// ==========================================================
{
    title:"ERROR_DETECTION",

    keywords:[
        "find excel errors",
        "error check",
        "formula error",
        "detect errors",
        "excel audit",
        "spreadsheet quality",
        "data validation"
    ],

    aliases:[
        "error checker",
        "formula monitoring",
        "excel health check"
    ],

    questions:[
        "How to identify formula errors?",
        "How to check worksheet errors?",
        "How to create error report?"
    ],

    dataset:"worksheet_audit",

    category:"AI Assistant",

    formula:'=IFERROR(A2,"Error Found")',

    syntax:"IFERROR(value,value_if_error)",

    explanation:"Detects formula errors and replaces them with a custom message.",

    breakdown:[
        "A2 = Formula Result",
        "Checks Error Condition",
        "Returns Status"
    ],

    output:"Error Found",

    business:"Used for automated spreadsheet quality checks.",

    tips:[
        "Combine with conditional formatting.",
        "Useful before publishing reports."
    ],

    errors:[
        "#VALUE!",
        "#REF!",
        "#DIV/0!"
    ],

    alternative:"=ISERROR()"
},


// ==========================================================
// REPORT_STATUS
// ==========================================================
{
    title:"REPORT_STATUS",

    keywords:[
        "report status",
        "mis status",
        "completed report",
        "pending report",
        "dashboard status",
        "automation status"
    ],

    aliases:[
        "report tracker",
        "completion status",
        "mis monitoring"
    ],

    questions:[
        "How to create report status?",
        "How to track completion?",
        "How to automate MIS status?"
    ],

    dataset:"mis_tracker",

    category:"AI Assistant",

    formula:'=IF(B2=C2,"Completed","Pending")',

    syntax:"IF(actual=target,status1,status2)",

    explanation:"Creates automated status based on completion criteria.",

    breakdown:[
        "B2 = Completed Count",
        "C2 = Expected Count",
        "Returns Status"
    ],

    output:"Completed",

    business:"Used for MIS tracking and operational dashboards.",

    tips:[
        "Can include SLA status.",
        "Useful for daily reporting."
    ],

    errors:[
        "#VALUE!"
    ],

    alternative:"=IFS()"
},


// ==========================================================
// KPI_INDICATOR
// ==========================================================
{
    title:"KPI_INDICATOR",

    keywords:[
        "kpi status",
        "performance indicator",
        "dashboard indicator",
        "target status",
        "green red status",
        "management dashboard"
    ],

    aliases:[
        "traffic light formula",
        "kpi flag",
        "performance signal"
    ],

    questions:[
        "How to create KPI indicator?",
        "How to show target status?",
        "How to create dashboard signals?"
    ],

    dataset:"kpi_dashboard",

    category:"AI Assistant",

    formula:'=IF(B2>=90%,"Green",IF(B2>=70%,"Yellow","Red"))',

    syntax:"IF(score,indicator1,indicator2)",

    explanation:"Creates a visual KPI indicator based on performance percentage.",

    breakdown:[
        "B2 = Achievement %",
        "Checks Threshold",
        "Returns Indicator"
    ],

    output:"Green",

    business:"Used in executive dashboards and performance reports.",

    tips:[
        "Combine with conditional formatting icons.",
        "Customize thresholds as required."
    ],

    errors:[
        "#VALUE!"
    ],

    alternative:"=IFS()"
},


// ==========================================================
// SMART_SEARCH_MAPPING
// ==========================================================
{
    title:"SMART_SEARCH_MAPPING",

    keywords:[
        "excel ai search",
        "formula search",
        "smart search",
        "natural language search",
        "find formula",
        "ai assistant",
        "formula database"
    ],

    aliases:[
        "keyword mapping",
        "search engine logic",
        "formula matching"
    ],

    questions:[
        "How to build formula search?",
        "How to match user query?",
        "How does AI formula search work?"
    ],

    dataset:"formula_keywords",

    category:"AI Assistant",

    formula:'=XLOOKUP(A2,Keywords,Formula,"Not Found")',

    syntax:"XLOOKUP(search_value,lookup_array,result_array)",

    explanation:"Maps user keywords to the correct Excel formula.",

    breakdown:[
        "A2 = User Query",
        "Keywords = Formula Database",
        "Returns Matching Formula"
    ],

    output:"=SUMIFS()",

    business:"Used as the backend logic of Excel AI assistants.",

    tips:[
        "Add synonyms for better matching.",
        "Use JavaScript fuzzy search for AI experience."
    ],

    errors:[
        "#N/A"
    ],

    alternative:"INDEX MATCH"
},

// ==========================================================
// QUERY_INTENT_DETECTION
// ==========================================================
{
    title:"QUERY_INTENT_DETECTION",

    keywords:[
        "understand user query",
        "detect requirement",
        "user intention",
        "natural language",
        "ai search",
        "formula recommendation"
    ],

    aliases:[
        "intent analyzer",
        "query understanding",
        "ai requirement detection"
    ],

    questions:[
        "How AI understands Excel requirement?",
        "How to detect user intent?"
    ],

    dataset:"user_queries",

    category:"AI Engine",

    formula:'=IF(ISNUMBER(SEARCH("total",A2)),"SUM",IF(ISNUMBER(SEARCH("lookup",A2)),"XLOOKUP","Search More"))',

    syntax:"SEARCH(keyword,text)",

    explanation:"Identifies user requirement keywords before recommending formulas.",

    breakdown:[
        "A2 = User Query",
        "SEARCH = Keyword Matching",
        "Returns Formula Category"
    ],

    output:"SUM",

    business:"Used as the first layer of an Excel AI assistant.",

    tips:[
        "Add synonyms for better accuracy.",
        "Use JavaScript for advanced matching."
    ],

    errors:[
        "#VALUE!"
    ],

    alternative:"AI NLP Model"
},


// ==========================================================
// FORMULA_MATCH_SCORE
// ==========================================================
{
    title:"FORMULA_MATCH_SCORE",

    keywords:[
        "formula ranking",
        "best formula",
        "formula score",
        "match percentage",
        "formula accuracy",
        "recommendation score"
    ],

    aliases:[
        "formula confidence",
        "matching score",
        "ai ranking"
    ],

    questions:[
        "How to rank formulas?",
        "How to select best formula?"
    ],

    dataset:"formula_database",

    category:"AI Engine",

    formula:"=COUNTIF(Keywords,A2)/COUNTA(Keywords)*100",

    syntax:"Matching Keywords / Total Keywords",

    explanation:"Calculates how closely a formula matches the user's requirement.",

    breakdown:[
        "A2 = User Keywords",
        "Keyword Database Matching",
        "Returns Match Score"
    ],

    output:"92%",

    business:"Used to rank formula suggestions.",

    tips:[
        "Higher score should appear first.",
        "Useful for AI recommendations."
    ],

    errors:[
        "#DIV/0!"
    ],

    alternative:"Fuzzy Search Algorithm"
},


// ==========================================================
// RELATED_FORMULA_SUGGESTION
// ==========================================================
{
    title:"RELATED_FORMULA_SUGGESTION",

    keywords:[
        "related formula",
        "alternative formula",
        "similar function",
        "suggest another formula",
        "formula recommendation"
    ],

    aliases:[
        "formula alternatives",
        "similar functions",
        "next best formula"
    ],

    questions:[
        "How to suggest related formulas?",
        "How to recommend alternatives?"
    ],

    dataset:"formula_relationship",

    category:"AI Engine",

    formula:'=XLOOKUP(A2,Formula_Name,Related_Formula)',

    syntax:"XLOOKUP(search,result)",

    explanation:"Provides alternative formulas related to the selected formula.",

    breakdown:[
        "A2 = Selected Formula",
        "Search Formula Database",
        "Return Related Formula"
    ],

    output:"SUMIFS → Pivot Table",

    business:"Helps users learn multiple Excel approaches.",

    tips:[
        "Add beginner and advanced alternatives.",
        "Improves learning experience."
    ],

    errors:[
        "#N/A"
    ],

    alternative:"INDEX MATCH"
},


// ==========================================================
// DIFFICULTY_LEVEL
// ==========================================================
{
    title:"DIFFICULTY_LEVEL",

    keywords:[
        "formula difficulty",
        "beginner formula",
        "advanced formula",
        "excel learning level",
        "skill level"
    ],

    aliases:[
        "formula complexity",
        "learning category",
        "excel level"
    ],

    questions:[
        "How to classify formulas?",
        "How to show learning level?"
    ],

    dataset:"formula_master",

    category:"AI Engine",

    formula:'=IF(B2<=2,"Beginner",IF(B2<=4,"Intermediate","Advanced"))',

    syntax:"IF(level,category)",

    explanation:"Classifies formulas according to complexity.",

    breakdown:[
        "B2 = Complexity Score",
        "Checks Level",
        "Returns Category"
    ],

    output:"Advanced",

    business:"Used for Excel learning platforms.",

    tips:[
        "Useful for your YouTube course integration.",
        "Helps users learn step-by-step."
    ],

    errors:[
        "#VALUE!"
    ],

    alternative:"IFS()"
},


// ==========================================================
// USE_CASE_MAPPING
// ==========================================================
{
    title:"USE_CASE_MAPPING",

    keywords:[
        "business use case",
        "formula application",
        "industry example",
        "real life example",
        "excel scenario"
    ],

    aliases:[
        "formula usage",
        "business mapping",
        "practical example"
    ],

    questions:[
        "How to show formula examples?",
        "How to map formulas with business?"
    ],

    dataset:"formula_use_cases",

    category:"AI Engine",

    formula:'=XLOOKUP(A2,Formula,Industry_Example)',

    syntax:"XLOOKUP(formula,use_case)",

    explanation:"Connects Excel formulas with practical business scenarios.",

    breakdown:[
        "A2 = Formula Name",
        "Search Use Case Database",
        "Returns Business Example"
    ],

    output:"XLOOKUP → Employee Master Lookup",

    business:"Makes Excel learning practical and job-oriented.",

    tips:[
        "Add Banking, HR, Finance examples.",
        "Differentiate your ExcelAI project."
    ],

    errors:[
        "#N/A"
    ],

    alternative:"Database Mapping"
},
// ==========================================================
// FORMULA_EXPLANATION_GENERATOR
// ==========================================================
{
    title:"FORMULA_EXPLANATION_GENERATOR",

    keywords:[
        "explain formula",
        "formula explanation",
        "what formula does",
        "understand excel formula",
        "learn excel",
        "excel tutor",
        "formula meaning"
    ],

    aliases:[
        "formula teacher",
        "formula guide",
        "excel explanation"
    ],

    questions:[
        "Explain this formula?",
        "What does this formula do?",
        "How to understand formula?"
    ],

    dataset:"formula_library",

    category:"AI Learning",

    formula:'=XLOOKUP(A2,Formula_Name,Explanation)',

    syntax:"XLOOKUP(search_value,formula_database,explanation)",

    explanation:"Retrieves a simple explanation of the selected Excel formula.",

    breakdown:[
        "A2 = Formula Name",
        "Search Formula Database",
        "Return Explanation"
    ],

    output:"XLOOKUP finds matching records from another table.",

    business:"Used for creating Excel learning assistants.",

    tips:[
        "Add beginner-friendly explanations.",
        "Avoid technical language only."
    ],

    errors:[
        "#N/A"
    ],

    alternative:"AI Explanation Model"
},


// ==========================================================
// STEP_BY_STEP_BREAKDOWN
// ==========================================================
{
    title:"STEP_BY_STEP_BREAKDOWN",

    keywords:[
        "formula steps",
        "step by step explanation",
        "formula breakdown",
        "learn formula",
        "excel training",
        "formula tutorial"
    ],

    aliases:[
        "formula steps",
        "calculation logic",
        "formula process"
    ],

    questions:[
        "Show formula steps?",
        "How formula works internally?"
    ],

    dataset:"formula_steps",

    category:"AI Learning",

    formula:'=TEXTJOIN(" → ",TRUE,A2:D2)',

    syntax:"TEXTJOIN(delimiter,ignore_empty,text)",

    explanation:"Combines multiple formula explanation steps into a learning sequence.",

    breakdown:[
        "A2:D2 = Formula Steps",
        "TEXTJOIN Combines Steps",
        "Creates Learning Flow"
    ],

    output:"Range → Condition → Calculation → Result",

    business:"Used for formula tutorials and training content.",

    tips:[
        "Useful for YouTube Excel lessons.",
        "Improves user understanding."
    ],

    errors:[
        "#VALUE!"
    ],

    alternative:"Manual Explanation"
},


// ==========================================================
// EXAMPLE_GENERATOR
// ==========================================================
{
    title:"EXAMPLE_GENERATOR",

    keywords:[
        "formula example",
        "real life example",
        "excel example",
        "formula practice",
        "business example",
        "sample calculation"
    ],

    aliases:[
        "formula demo",
        "practical example",
        "formula scenario"
    ],

    questions:[
        "Show formula example?",
        "Give practical example?"
    ],

    dataset:"formula_examples",

    category:"AI Learning",

    formula:'=XLOOKUP(A2,Formula_Name,Example)',

    syntax:"XLOOKUP(formula,example_database)",

    explanation:"Provides practical examples for Excel formulas.",

    breakdown:[
        "A2 = Formula Name",
        "Search Example Database",
        "Return Practical Example"
    ],

    output:"SUMIFS → Calculate department salary total",

    business:"Used for Excel training platforms.",

    tips:[
        "Add Banking, HR, Finance examples.",
        "Keep examples job-oriented."
    ],

    errors:[
        "#N/A"
    ],

    alternative:"Example Library"
},


// ==========================================================
// DUMMY_DATA_GENERATOR
// ==========================================================
{
    title:"DUMMY_DATA_GENERATOR",

    keywords:[
        "dummy data",
        "sample data",
        "practice data",
        "excel dataset",
        "test data",
        "training data"
    ],

    aliases:[
        "sample table generator",
        "practice dataset",
        "excel data creator"
    ],

    questions:[
        "Create sample data?",
        "Generate practice dataset?"
    ],

    dataset:"sample_data",

    category:"AI Learning",

    formula:'=RANDBETWEEN(1000,5000)',

    syntax:"RANDBETWEEN(bottom,top)",

    explanation:"Generates sample values for Excel practice and testing.",

    breakdown:[
        "1000 = Minimum Value",
        "5000 = Maximum Value",
        "Returns Random Number"
    ],

    output:"3500",

    business:"Used for demonstrations, training, and testing formulas.",

    tips:[
        "Use realistic business data.",
        "Avoid personal information."
    ],

    errors:[
        "#NUM!"
    ],

    alternative:"Power Query Sample Data"
},


// ==========================================================
// COMMON_MISTAKE_DETECTOR
// ==========================================================
{
    title:"COMMON_MISTAKE_DETECTOR",

    keywords:[
        "formula mistake",
        "excel error reason",
        "wrong formula",
        "formula troubleshooting",
        "excel help",
        "debug formula"
    ],

    aliases:[
        "formula debugger",
        "error explanation",
        "excel troubleshooting"
    ],

    questions:[
        "Why formula not working?",
        "Find formula mistake?",
        "Explain error?"
    ],

    dataset:"excel_errors",

    category:"AI Learning",

    formula:'=IFERROR(A2,"Check Formula Logic")',

    syntax:"IFERROR(value,error_message)",

    explanation:"Identifies possible formula problems and provides guidance.",

    breakdown:[
        "A2 = Formula Result",
        "Check Error",
        "Return Guidance"
    ],

    output:"Check Formula Logic",

    business:"Used for Excel support and troubleshooting assistants.",

    tips:[
        "Maintain error knowledge database.",
        "Add examples of common mistakes."
    ],

    errors:[
        "#VALUE!",
        "#REF!",
        "#DIV/0!"
    ],

    alternative:"AI Error Analyzer"
},

// ==========================================================
// RESUME_PROJECT_FORMULA_MAPPING
// ==========================================================
{
    title:"RESUME_PROJECT_FORMULA_MAPPING",

    keywords:[
        "resume project",
        "mention excel project",
        "resume formula",
        "project description",
        "excel automation project",
        "career support",
        "data analyst resume"
    ],

    aliases:[
        "resume builder",
        "project suggestion",
        "cv excel project"
    ],

    questions:[
        "How to describe Excel project in resume?",
        "How to add Excel automation project?"
    ],

    dataset:"resume_projects",

    category:"Career Assistant",

    formula:'=XLOOKUP(A2,Skill,Project_Description)',

    syntax:"XLOOKUP(skill,project_database,description)",

    explanation:"Maps Excel skills with professional resume project descriptions.",

    breakdown:[
        "A2 = Excel Skill",
        "Search Project Database",
        "Return Resume Description"
    ],

    output:"Automated MIS Dashboard using Power Query and Excel",

    business:"Helps professionals create stronger resumes.",

    tips:[
        "Add business impact numbers.",
        "Mention automation tools used."
    ],

    errors:[
        "#N/A"
    ],

    alternative:"AI Resume Generator"
},


// ==========================================================
// BANKING_MIS_FORMULA_SUGGESTION
// ==========================================================
{
    title:"BANKING_MIS_FORMULA_SUGGESTION",

    keywords:[
        "banking mis",
        "treasury report",
        "compliance report",
        "sla report",
        "risk report",
        "financial dashboard",
        "bank excel"
    ],

    aliases:[
        "banking formula",
        "treasury analytics",
        "compliance automation"
    ],

    questions:[
        "Which formula for banking MIS?",
        "How to automate treasury reports?"
    ],

    dataset:"banking_mis",

    category:"Professional MIS",

    formula:'=XLOOKUP(A2,MIS_Requirement,Formula)',

    syntax:"XLOOKUP(requirement,formula_database)",

    explanation:"Suggests formulas based on banking and MIS requirements.",

    breakdown:[
        "A2 = MIS Requirement",
        "Match Business Scenario",
        "Return Formula"
    ],

    output:"Ageing Report → DATEDIF",

    business:"Used for treasury, compliance and operational reporting.",

    tips:[
        "Add RBI, SOX, ICOFR scenarios.",
        "Include real business examples."
    ],

    errors:[
        "#N/A"
    ],

    alternative:"Power BI Measures"
},


// ==========================================================
// INTERVIEW_QUESTION_GENERATOR
// ==========================================================
{
    title:"INTERVIEW_QUESTION_GENERATOR",

    keywords:[
        "excel interview",
        "excel questions",
        "data analyst interview",
        "formula interview",
        "practice interview",
        "job preparation"
    ],

    aliases:[
        "excel quiz",
        "interview preparation",
        "excel assessment"
    ],

    questions:[
        "Generate Excel interview questions?",
        "Prepare for Excel interview?"
    ],

    dataset:"excel_interview_questions",

    category:"Career Assistant",

    formula:'=XLOOKUP(A2,Formula,Interview_Question)',

    syntax:"XLOOKUP(formula,question_database)",

    explanation:"Provides interview questions related to Excel formulas.",

    breakdown:[
        "A2 = Formula Name",
        "Search Question Database",
        "Return Interview Question"
    ],

    output:"Explain difference between VLOOKUP and XLOOKUP",

    business:"Helps users prepare for Excel interviews.",

    tips:[
        "Add scenario-based questions.",
        "Include practical tasks."
    ],

    errors:[
        "#N/A"
    ],

    alternative:"AI Question Generator"
},


// ==========================================================
// EXCEL_SKILL_ASSESSMENT
// ==========================================================
{
    title:"EXCEL_SKILL_ASSESSMENT",

    keywords:[
        "excel skill test",
        "excel assessment",
        "skill level",
        "beginner test",
        "advanced excel test",
        "learning evaluation"
    ],

    aliases:[
        "excel exam",
        "knowledge test",
        "skill evaluation"
    ],

    questions:[
        "How to test Excel skills?",
        "Create Excel assessment?"
    ],

    dataset:"skill_test",

    category:"Career Assistant",

    formula:'=IF(B2>=80%,"Advanced",IF(B2>=50%,"Intermediate","Beginner"))',

    syntax:"IF(score,skill_level)",

    explanation:"Classifies Excel skill level based on assessment score.",

    breakdown:[
        "B2 = Test Score",
        "Compare Threshold",
        "Return Skill Level"
    ],

    output:"Advanced",

    business:"Used for training and employee skill evaluation.",

    tips:[
        "Create formula-wise questions.",
        "Track improvement over time."
    ],

    errors:[
        "#VALUE!"
    ],

    alternative:"IFS()"
},


// ==========================================================
// FORMULA_PRACTICE_CHALLENGE
// ==========================================================
{
    title:"FORMULA_PRACTICE_CHALLENGE",

    keywords:[
        "formula practice",
        "excel challenge",
        "learn formula",
        "practice exercise",
        "excel task",
        "hands on learning"
    ],

    aliases:[
        "excel exercise",
        "formula challenge",
        "practice test"
    ],

    questions:[
        "Give Excel practice task?",
        "Create formula challenge?"
    ],

    dataset:"excel_challenges",

    category:"Career Assistant",

    formula:'=XLOOKUP(A2,Formula,Practice_Task)',

    syntax:"XLOOKUP(formula,challenge_database)",

    explanation:"Creates practical Excel exercises for learners.",

    breakdown:[
        "A2 = Formula",
        "Find Related Challenge",
        "Return Practice Task"
    ],

    output:"Create employee ageing report using IF",

    business:"Used for Excel learning platforms and courses.",

    tips:[
        "Add difficulty levels.",
        "Include real business datasets."
    ],

    errors:[
        "#N/A"
    ],

    alternative:"Learning Management System"
},

// ==========================================================
// POWER_QUERY_AUTOMATION_SUGGESTION
// ==========================================================
{
    title:"POWER_QUERY_AUTOMATION_SUGGESTION",

    keywords:[
        "power query",
        "merge excel files",
        "append files",
        "etl automation",
        "data transformation",
        "excel automation",
        "monthly report automation"
    ],

    aliases:[
        "query automation",
        "excel etl",
        "data cleaning automation"
    ],

    questions:[
        "How to automate Excel files?",
        "How to merge monthly reports?"
    ],

    dataset:"automation_library",

    category:"Automation",

    formula:'=XLOOKUP(A2,Requirement,Power_Query_Solution)',

    syntax:"XLOOKUP(requirement,automation_database)",

    explanation:"Suggests Power Query solutions based on automation requirements.",

    breakdown:[
        "A2 = User Requirement",
        "Search Automation Database",
        "Return Power Query Solution"
    ],

    output:"Append Multiple Excel Files",

    business:"Used for ETL automation and recurring MIS reports.",

    tips:[
        "Use Power Query for repeatable processes.",
        "Reduce manual copy-paste work."
    ],

    errors:[
        "#N/A"
    ],

    alternative:"Power Pivot"
},


// ==========================================================
// POWER_BI_MAPPING
// ==========================================================
{
    title:"POWER_BI_MAPPING",

    keywords:[
        "power bi",
        "dashboard",
        "visualization",
        "kpi dashboard",
        "data model",
        "business intelligence"
    ],

    aliases:[
        "bi dashboard",
        "report visualization",
        "power bi suggestion"
    ],

    questions:[
        "Which Power BI visual to use?",
        "How to create dashboard?"
    ],

    dataset:"powerbi_templates",

    category:"Business Intelligence",

    formula:'=XLOOKUP(A2,KPI,Recommended_Visual)',

    syntax:"XLOOKUP(kpi,visual_database)",

    explanation:"Maps business KPIs with suitable Power BI visuals.",

    breakdown:[
        "A2 = KPI Name",
        "Search Visual Library",
        "Return Recommended Chart"
    ],

    output:"Sales Trend → Line Chart",

    business:"Used for executive dashboards and reporting.",

    tips:[
        "Select visuals based on business questions.",
        "Avoid unnecessary charts."
    ],

    errors:[
        "#N/A"
    ],

    alternative:"Power BI Recommendation Engine"
},


// ==========================================================
// SQL_QUERY_RECOMMENDATION
// ==========================================================
{
    title:"SQL_QUERY_RECOMMENDATION",

    keywords:[
        "sql query",
        "database query",
        "select data",
        "join tables",
        "sql automation",
        "data extraction"
    ],

    aliases:[
        "sql suggestion",
        "database formula",
        "query generator"
    ],

    questions:[
        "Generate SQL query?",
        "How to retrieve data?"
    ],

    dataset:"sql_library",

    category:"Data Analytics",

    formula:'=XLOOKUP(A2,Requirement,SQL_Query)',

    syntax:"XLOOKUP(requirement,SQL_database)",

    explanation:"Suggests SQL queries based on reporting requirements.",

    breakdown:[
        "A2 = Business Requirement",
        "Match SQL Pattern",
        "Return Query"
    ],

    output:"Employee Report → SELECT query",

    business:"Used for database reporting and analytics.",

    tips:[
        "Add SQL examples by business scenario.",
        "Include joins and aggregations."
    ],

    errors:[
        "#N/A"
    ],

    alternative:"AI SQL Generator"
},


// ==========================================================
// PYTHON_AUTOMATION_SUGGESTION
// ==========================================================
{
    title:"PYTHON_AUTOMATION_SUGGESTION",

    keywords:[
        "python automation",
        "selenium",
        "python script",
        "data automation",
        "file processing",
        "web scraping"
    ],

    aliases:[
        "python helper",
        "automation script",
        "python recommendation"
    ],

    questions:[
        "Where to use Python automation?",
        "Suggest Python solution?"
    ],

    dataset:"python_use_cases",

    category:"Automation",

    formula:'=XLOOKUP(A2,Task,Python_Solution)',

    syntax:"XLOOKUP(task,python_database)",

    explanation:"Maps business tasks with Python automation solutions.",

    breakdown:[
        "A2 = Task Requirement",
        "Search Python Library",
        "Return Automation Approach"
    ],

    output:"Multiple PDF Extraction → Python Script",

    business:"Used for advanced data automation.",

    tips:[
        "Use Python for repetitive tasks.",
        "Combine with Excel workflows."
    ],

    errors:[
        "#N/A"
    ],

    alternative:"Power Automate"
},


// ==========================================================
// POWER_AUTOMATE_WORKFLOW_MAPPING
// ==========================================================
{
    title:"POWER_AUTOMATE_WORKFLOW_MAPPING",

    keywords:[
        "power automate",
        "workflow",
        "approval automation",
        "email automation",
        "notification",
        "office automation"
    ],

    aliases:[
        "workflow builder",
        "automation flow",
        "process automation"
    ],

    questions:[
        "How to automate workflow?",
        "Create approval flow?"
    ],

    dataset:"workflow_library",

    category:"Automation",

    formula:'=XLOOKUP(A2,Process,Flow_Design)',

    syntax:"XLOOKUP(process,workflow_database)",

    explanation:"Suggests Power Automate workflows for business processes.",

    breakdown:[
        "A2 = Business Process",
        "Find Workflow Pattern",
        "Return Automation Flow"
    ],

    output:"Leave Approval → Approval Workflow",

    business:"Used for business process automation.",

    tips:[
        "Connect Excel, Outlook, SharePoint.",
        "Track automation savings."
    ],

    errors:[
        "#N/A"
    ],

    alternative:"Custom API Automation"
},

// ==========================================================
// FORMULA_COMPARISON
// ==========================================================
{
    title:"FORMULA_COMPARISON",

    keywords:[
        "compare formulas",
        "formula comparison",
        "difference between formulas",
        "which formula is better",
        "best excel formula",
        "alternative formula",
        "excel comparison"
    ],

    aliases:[
        "formula difference",
        "compare excel functions",
        "formula analysis"
    ],

    questions:[
        "Compare Excel formulas?",
        "Which formula is better?",
        "Show formula alternatives?"
    ],

    dataset:"formula_comparison",

    category:"AI Engine",

    formula:'=XLOOKUP(A2,Formula_Name,Comparison_Data)',

    syntax:"XLOOKUP(formula,comparison_database)",

    explanation:"Compares similar Excel formulas and shows differences.",

    breakdown:[
        "A2 = Formula Name",
        "Search Comparison Database",
        "Return Comparison Result"
    ],

    output:"XLOOKUP vs VLOOKUP",

    business:"Used for Excel learning and formula selection.",

    tips:[
        "Show advantages and limitations.",
        "Recommend modern formulas first."
    ],

    errors:[
        "#N/A"
    ],

    alternative:"AI Comparison Engine"
},


// ==========================================================
// FORMULA_ADVANTAGE_ANALYSIS
// ==========================================================
{
    title:"FORMULA_ADVANTAGE_ANALYSIS",

    keywords:[
        "formula advantage",
        "formula benefit",
        "why use xlookup",
        "formula pros cons",
        "excel improvement"
    ],

    aliases:[
        "formula benefit",
        "formula strength",
        "formula evaluation"
    ],

    questions:[
        "Why use this formula?",
        "What are formula benefits?"
    ],

    dataset:"formula_features",

    category:"AI Engine",

    formula:'=TEXTJOIN(", ",TRUE,A2:D2)',

    syntax:"TEXTJOIN(separator,values)",

    explanation:"Combines formula advantages into a readable explanation.",

    breakdown:[
        "A2:D2 = Feature List",
        "TEXTJOIN Combines Points",
        "Returns Explanation"
    ],

    output:"Dynamic, Faster, Easy Maintenance",

    business:"Helps users understand formula selection.",

    tips:[
        "Include limitations also.",
        "Avoid recommending blindly."
    ],

    errors:[
        "#VALUE!"
    ],

    alternative:"AI Explanation"
},


// ==========================================================
// VERSION_COMPATIBILITY
// ==========================================================
{
    title:"VERSION_COMPATIBILITY",

    keywords:[
        "excel version",
        "365 formula",
        "old excel",
        "compatibility",
        "supported version",
        "formula availability"
    ],

    aliases:[
        "excel support",
        "formula version check",
        "office compatibility"
    ],

    questions:[
        "Does this formula work in my Excel?",
        "Which Excel version supports this?"
    ],

    dataset:"excel_versions",

    category:"AI Engine",

    formula:'=XLOOKUP(A2,Formula,Excel_Version)',

    syntax:"XLOOKUP(formula,version_database)",

    explanation:"Shows Excel versions supported by formulas.",

    breakdown:[
        "A2 = Formula Name",
        "Check Version Database",
        "Return Compatibility"
    ],

    output:"XLOOKUP → Excel 365 / 2021",

    business:"Helps users select compatible formulas.",

    tips:[
        "Important for corporate users.",
        "Mention legacy alternatives."
    ],

    errors:[
        "#N/A"
    ],

    alternative:"INDEX MATCH"
},


// ==========================================================
// PERFORMANCE_RANKING
// ==========================================================
{
    title:"PERFORMANCE_RANKING",

    keywords:[
        "formula speed",
        "performance comparison",
        "fast formula",
        "excel optimization",
        "large data performance"
    ],

    aliases:[
        "formula efficiency",
        "formula optimization",
        "calculation speed"
    ],

    questions:[
        "Which formula is faster?",
        "Compare formula performance?"
    ],

    dataset:"formula_performance",

    category:"AI Engine",

    formula:'=RANK.EQ(B2,$B$2:$B$10,0)',

    syntax:"RANK.EQ(value,range)",

    explanation:"Ranks formulas based on performance score.",

    breakdown:[
        "B2 = Formula Performance Score",
        "Compare Formula Scores",
        "Return Ranking"
    ],

    output:"1",

    business:"Useful for optimizing large Excel models.",

    tips:[
        "Avoid volatile functions where possible.",
        "Consider file size impact."
    ],

    errors:[
        "#N/A"
    ],

    alternative:"Manual Benchmark"
},


// ==========================================================
// BEST_FORMULA_RECOMMENDATION
// ==========================================================
{
    title:"BEST_FORMULA_RECOMMENDATION",

    keywords:[
        "best formula",
        "recommended formula",
        "choose formula",
        "formula suggestion",
        "excel ai recommendation"
    ],

    aliases:[
        "formula selector",
        "smart recommendation",
        "formula decision engine"
    ],

    questions:[
        "Which formula should I use?",
        "Recommend best formula?"
    ],

    dataset:"formula_decision_engine",

    category:"AI Engine",

    formula:'=INDEX(A2:A10,MATCH(MAX(B2:B10),B2:B10,0))',

    syntax:"INDEX + MATCH + MAX",

    explanation:"Selects the highest-ranked formula based on scoring criteria.",

    breakdown:[
        "A Column = Formula Names",
        "B Column = Formula Score",
        "MAX Finds Best Score",
        "INDEX Returns Formula"
    ],

    output:"XLOOKUP",

    business:"Core logic for Excel AI recommendation systems.",

    tips:[
        "Consider user Excel version.",
        "Consider data size and complexity."
    ],

    errors:[
        "#N/A"
    ],

    alternative:"AI Ranking Model"
}

];


/*======================================
 Search Formula
======================================*/

function searchFormula(question){

    question = question.toLowerCase();

    for(let item of AI_DATABASE){

        for(let key of item.keywords){

            if(question.includes(key)){

                return item;

            }

        }

    }

    return null;

}

console.log("✅ ai.js Loaded");
