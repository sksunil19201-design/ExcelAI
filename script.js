// ==============================
// ExcelAI - script.js
// ==============================

// Elements
const problem = document.getElementById("problem");
const generateBtn = document.getElementById("generateBtn");
const formula = document.getElementById("formula");
const explanation = document.getElementById("explanation");
const example = document.getElementById("example");
const copyBtn = document.getElementById("copyBtn");

// =======================================
// Formula Database
// =======================================

const FORMULAS = [

{
title:"XLOOKUP",
keywords:[
"xlookup","lookup","employee id","lookup employee","search employee",
"find employee","employee name","employee details","lookup id",
"find record","fetch employee","search id","return value","lookup data","office 365 lookup"
],
formula:'=XLOOKUP(A2,Sheet2!A:A,Sheet2!B:B,"Not Found")',
explanation:"Returns employee name from Employee ID.",
example:"10025 → Rahul Sharma"
},

{
title:"XLOOKUP Salary",
keywords:[
"xlookup salary","salary lookup","employee salary","salary by id","find salary",
"lookup salary","fetch salary","salary details","employee pay",
"salary search","salary data","xlookup employee salary"
],
formula:'=XLOOKUP(A2,Sheet2!A:A,Sheet2!C:C)',
explanation:"Returns employee salary.",
example:"10025 → ₹35,000"
},

{
title:"XLOOKUP Department",
keywords:[
"xlookup department","department lookup","employee department","find department",
"lookup department","department by id","employee team","hr department",
"fetch department","department search","office department"
],
formula:'=XLOOKUP(A2,Sheet2!A:A,Sheet2!D:D)',
explanation:"Returns employee department.",
example:"10025 → HR"
},

{
title:"VLOOKUP",
keywords:[
"vlookup","lookup","vertical lookup","find value","search value",
"employee lookup","lookup employee","retrieve data","lookup data",
"fetch data","excel lookup","find employee","table lookup","search record","v lookup"
],
formula:'=VLOOKUP(A2,Sheet2!A:D,2,FALSE)',
explanation:"Looks up a matching value vertically.",
example:"Employee Name"
},

{
title:"HLOOKUP",
keywords:[
"hlookup","horizontal lookup","lookup row","row lookup","find value row",
"table lookup","header lookup","horizontal search","excel hlookup",
"lookup horizontally"
],
formula:'=HLOOKUP(A2,A1:H5,2,FALSE)',
explanation:"Looks up a matching value horizontally.",
example:"Monthly Sales"
},

{
title:"INDEX MATCH",
keywords:[
"index match","match index","lookup alternative","advanced lookup",
"index","match","search value","find record","lookup row",
"dynamic lookup","flexible lookup","excel lookup","index formula","match formula","lookup without vlookup"
],
formula:'=INDEX(B:B,MATCH(A2,A:A,0))',
explanation:"Flexible lookup using INDEX and MATCH.",
example:"Employee Name"
},

{
title:"MATCH",
keywords:[
"match","find position","row position","lookup position","search position",
"find row","row number","match value","excel match","position formula"
],
formula:'=MATCH(A2,A:A,0)',
explanation:"Returns the position of a value.",
example:"Position 15"
},

{
title:"INDEX",
keywords:[
"index","return value","value by row","lookup value","cell value",
"index formula","retrieve value","excel index","find data","reference value"
],
formula:'=INDEX(B:B,5)',
explanation:"Returns a value from a specified row.",
example:"Rahul"
},

{
title:"CHOOSE",
keywords:[
"choose","select value","choose option","pick value","index selection",
"multiple choice","excel choose","return option","choose formula","selection"
],
formula:'=CHOOSE(2,"HR","Finance","IT")',
explanation:"Returns a value based on its index.",
example:"Finance"
},

{
title:"OFFSET",
keywords:[
"offset","dynamic reference","offset cell","move reference","relative reference",
"cell offset","excel offset","dynamic range","shift cell","offset formula"
],
formula:'=OFFSET(A1,2,1)',
explanation:"Returns a reference offset from a starting cell.",
example:"Cell B3"
},

{
title:"INDIRECT",
keywords:[
"indirect","dynamic reference","text reference","cell reference",
"reference text","excel indirect","dynamic cell","reference formula",
"indirect formula","convert text reference"
],
formula:'=INDIRECT("A2")',
explanation:"Returns a reference from text.",
example:"Value from A2"
},

{
title:"ADDRESS",
keywords:[
"address","cell address","cell reference","address formula","row column",
"reference address","excel address","cell location","address function","reference"
],
formula:'=ADDRESS(5,2)',
explanation:"Returns the address of a cell.",
example:"$B$5"
},

{
title:"ROW",
keywords:[
"row","row number","current row","excel row","find row",
"row formula","row position","row reference","row index","row function"
],
formula:'=ROW(A5)',
explanation:"Returns the row number.",
example:"5"
},

{
title:"COLUMN",
keywords:[
"column","column number","current column","excel column","find column",
"column formula","column position","column index","column function","column reference"
],
formula:'=COLUMN(C1)',
explanation:"Returns the column number.",
example:"3"
},

{
title:"ROWS",
keywords:[
"rows","count rows","number of rows","row count","total rows",
"excel rows","rows formula","count records","row total","rows function"
],
formula:'=ROWS(A1:A20)',
explanation:"Counts the total rows.",
example:"20"
},

{
title:"COLUMNS",
keywords:[
"columns","count columns","number of columns","column count",
"total columns","excel columns","columns formula","column total",
"count fields","columns function"
],
formula:'=COLUMNS(A:F)',
explanation:"Counts the total columns.",
example:"6"
},

{
title:"FILTER",
keywords:[
"filter","filter data","filter rows","show matching records",
"extract records","dynamic filter","filter table","show employees",
"filter department","excel filter","filtered data","display records"
],
formula:'=FILTER(A2:D100,C2:C100="HR")',
explanation:"Returns only matching records.",
example:"Only HR Employees"
},

{
title:"UNIQUE",
keywords:[
"unique","remove duplicates","distinct values","unique list",
"unique records","distinct","unique formula","excel unique",
"duplicate removal","list unique"
],
formula:'=UNIQUE(A2:A100)',
explanation:"Returns unique values.",
example:"Unique Departments"
},

{
title:"SORT",
keywords:[
"sort","sort ascending","sort data","order data","ascending order",
"descending order","sort table","excel sort","sort records","sort list"
],
formula:'=SORT(A2:D100)',
explanation:"Sorts data in ascending order.",
example:"Employee List"
},

{
title:"SORTBY",
keywords:[
"sortby","sort by column","sort salary","sort department",
"sort descending","custom sort","sort using another column",
"excel sortby","dynamic sort","salary descending"
],
formula:'=SORTBY(A2:D100,D2:D100,-1)',
explanation:"Sorts using another column.",
example:"Highest Salary First"
},

{
title:"SEQUENCE",
keywords:[
"sequence","number sequence","generate numbers","serial numbers",
"auto numbers","number list","excel sequence","dynamic numbers",
"running numbers","generate series"
],
formula:'=SEQUENCE(10)',
explanation:"Generates a sequence of numbers.",
example:"1 to 10"
},

{
title:"TRANSPOSE",
keywords:[
"transpose","rows to columns","columns to rows","flip table",
"rotate data","transpose data","excel transpose","change orientation",
"vertical to horizontal","horizontal to vertical"
],
formula:'=TRANSPOSE(A1:D2)',
explanation:"Converts rows into columns.",
example:"Horizontal to Vertical"
},

{
title:"LEFT",
keywords:[
"left","first characters","extract left","left text","first letters",
"extract beginning","left function","text extraction","excel left","substring left"
],
formula:'=LEFT(A2,5)',
explanation:"Returns characters from the left.",
example:"Rahul"
},

{
title:"RIGHT",
keywords:[
"right","last characters","extract right","right text","last letters",
"extract ending","right function","text extraction","excel right","substring right"
],
formula:'=RIGHT(A2,4)',
explanation:"Returns characters from the right.",
example:"0025"
},

{
title:"MID",
keywords:[
"mid","extract middle","middle text","substring","text extract",
"mid function","excel mid","middle characters","extract text","partial text"
],
formula:'=MID(A2,3,5)',
explanation:"Extracts characters from the middle.",
example:"CDEFG"
},

{
title:"LEN",
keywords:[
"len","length","text length","count characters","character count",
"string length","excel len","length formula","text size","count letters"
],
formula:'=LEN(A2)',
explanation:"Returns the number of characters.",
example:"Rahul → 5"
},

{
title:"TRIM",
keywords:[
"trim","remove spaces","extra spaces","clean spaces","space remove",
"trim text","leading spaces","trailing spaces","text cleanup","excel trim"
],
formula:'=TRIM(A2)',
explanation:"Removes extra spaces from text.",
example:" Rahul Sharma  → Rahul Sharma"
},

{
title:"UPPER",
keywords:[
"upper","uppercase","capital letters","convert uppercase","all caps",
"text uppercase","upper function","excel upper","capital text","convert text"
],
formula:'=UPPER(A2)',
explanation:"Converts text to uppercase.",
example:"rahul → RAHUL"
},

{
title:"LOWER",
keywords:[
"lower","lowercase","small letters","convert lowercase","lower text",
"text lowercase","excel lower","lower function","convert text","small text"
],
formula:'=LOWER(A2)',
explanation:"Converts text to lowercase.",
example:"RAHUL → rahul"
},

{
title:"PROPER",
keywords:[
"proper","capitalize","title case","first letter capital","capitalize words",
"proper case","text formatting","excel proper","name format","capitalize text"
],
formula:'=PROPER(A2)',
explanation:"Capitalizes the first letter of each word.",
example:"rahul sharma → Rahul Sharma"
},

{
title:"CONCAT",
keywords:[
"concat","concatenate","join text","merge text","combine text",
"combine cells","join cells","text combine","merge columns","excel concat"
],
formula:'=CONCAT(A2," ",B2)',
explanation:"Combines text from multiple cells.",
example:"Rahul + Sharma"
},

{
title:"TEXTJOIN",
keywords:[
"textjoin","join text","combine text","merge cells","join multiple cells",
"concatenate range","join values","combine range","text combine","excel textjoin"
],
formula:'=TEXTJOIN(", ",TRUE,A2:A5)',
explanation:"Joins text using a delimiter.",
example:"HR, IT, Finance"
},

{
title:"FIND",
keywords:[
"find","find position","character position","locate text","find character",
"text position","find formula","excel find","locate character","search exact"
],
formula:'=FIND("@",A2)',
explanation:"Finds the position of a character.",
example:"Email @ position"
},

{
title:"SEARCH",
keywords:[
"search","find text","search text","locate word","case insensitive search",
"text search","excel search","find keyword","search formula","locate text"
],
formula:'=SEARCH("HR",A2)',
explanation:"Searches text without case sensitivity.",
example:"HR Department"
},

{
title:"REPLACE",
keywords:[
"replace","replace text","replace characters","change text","overwrite text",
"text replace","replace formula","excel replace","modify text","replace letters"
],
formula:'=REPLACE(A2,1,3,"EMP")',
explanation:"Replaces characters in text.",
example:"12345 → EMP45"
},

{
title:"SUBSTITUTE",
keywords:[
"substitute","replace word","change word","replace text","swap text",
"text substitute","replace string","excel substitute","modify word","find replace"
],
formula:'=SUBSTITUTE(A2,"HR","Human Resource")',
explanation:"Replaces matching text.",
example:"HR → Human Resource"
},

{
title:"TEXTBEFORE",
keywords:[
"textbefore","before delimiter","text before","extract before","left of delimiter",
"before dash","text extraction","excel textbefore","delimiter","split text"
],
formula:'=TEXTBEFORE(A2,"-")',
explanation:"Returns text before a delimiter.",
example:"EMP-1001 → EMP"
},

{
title:"TEXTAFTER",
keywords:[
"textafter","after delimiter","text after","extract after","right of delimiter",
"after dash","text extraction","excel textafter","delimiter","split text"
],
formula:'=TEXTAFTER(A2,"-")',
explanation:"Returns text after a delimiter.",
example:"EMP-1001 → 1001"
},

{
title:"TEXTSPLIT",
keywords:[
"textsplit","split text","split column","split cells","delimiter",
"text to columns","separate text","excel textsplit","split string","text separator"
],
formula:'=TEXTSPLIT(A2,",")',
explanation:"Splits text into multiple columns.",
example:"HR,IT,Finance"
},

{
title:"EXACT",
keywords:[
"exact","compare text","exact match","case sensitive","same text",
"text compare","excel exact","match text","equal text","exact comparison"
],
formula:'=EXACT(A2,B2)',
explanation:"Checks whether two text values are exactly the same.",
example:"TRUE/FALSE"
},

{
title:"REPT",
keywords:[
"rept","repeat text","repeat characters","duplicate text","repeat symbol",
"excel rept","repeat formula","text repeat","stars","mask text"
],
formula:'=REPT("*",5)',
explanation:"Repeats text a specified number of times.",
example:"*****"
},

{
title:"CLEAN",
keywords:[
"clean","remove non printable","clean data","remove junk","imported data",
"text cleanup","excel clean","clean formula","remove hidden characters","cleanup"
],
formula:'=CLEAN(A2)',
explanation:"Removes non-printable characters.",
example:"Clean imported data"
},

{
title:"TEXT",
keywords:[
"text","format date","format number","convert to text","display format",
"text formatting","excel text","date format","currency format","number format"
],
formula:'=TEXT(A2,"dd-mmm-yyyy")',
explanation:"Formats values as text.",
example:"01-Jan-2026"
},

{
title:"VALUE",
keywords:[
"value","text to number","convert number","numeric value","number conversion",
"convert text","excel value","text number","number format","convert numeric"
],
formula:'=VALUE(A2)',
explanation:"Converts text into a numeric value.",
example:"'100 → 100"
},

{
title:"DOLLAR",
keywords:[
"dollar","currency","money format","currency format","financial format",
"excel dollar","number currency","format amount","currency display","money"
],
formula:'=DOLLAR(A2,2)',
explanation:"Formats a number as currency.",
example:"$12,500.00"
},

{
title:"FIXED",
keywords:[
"fixed","fixed decimals","decimal format","number format","fixed number",
"excel fixed","round display","formatted number","fixed formula","display decimals"
],
formula:'=FIXED(A2,2)',
explanation:"Formats a number with fixed decimal places.",
example:"1234.57"
},

{
title:"CHAR",
keywords:[
"char","ascii","unicode","character code","ascii character",
"character","excel char","code to character","char function","text code"
],
formula:'=CHAR(65)',
explanation:"Returns the character represented by a code.",
example:"65 → A"
},

{
title:"CODE",
keywords:[
"code","ascii code","character code","unicode value","text code",
"excel code","character number","letter code","code function","ascii value"
],
formula:'=CODE(A2)',
explanation:"Returns the numeric code of the first character.",
example:"A → 65"
},

{
title:"UNICHAR",
keywords:[
"unichar","unicode character","special character","emoji","unicode text",
"character code","excel unichar","symbol","unicode symbol","special symbols"
],
formula:'=UNICHAR(9733)',
explanation:"Returns a Unicode character from a code.",
example:"★"
},

{
title:"UNICODE",
keywords:[
"unicode","unicode value","character number","unicode code","text code",
"excel unicode","symbol code","unicode function","character value","emoji code"
],
formula:'=UNICODE(A2)',
explanation:"Returns the Unicode number of the first character.",
example:"A → 65"
},

{
title:"IF",
keywords:[
"if","condition","logical test","if condition","pass fail",
"compare values","decision","if statement","check condition","true false",
"logical formula","conditional formula","if function","excel if","condition check"
],
formula:'=IF(A2>=50,"Pass","Fail")',
explanation:"Checks whether a condition is TRUE or FALSE.",
example:"70 → Pass"
},

{
title:"IFS",
keywords:[
"ifs","multiple if","multiple conditions","if else","nested if",
"grade calculation","logical conditions","multiple tests","excel ifs",
"condition formula","ifs function","multiple logical test","if formula","grading","conditions"
],
formula:'=IFS(A2>=90,"A",A2>=75,"B",A2>=50,"C",TRUE,"Fail")',
explanation:"Evaluates multiple conditions.",
example:"82 → B"
},

{
title:"IFERROR",
keywords:[
"iferror","error handling","hide error","remove error","excel error",
"divide by zero","replace error","if error","error formula","formula error",
"error message","handle errors","fix error","iferror function","error value"
],
formula:'=IFERROR(A2/B2,"Error")',
explanation:"Returns a custom value when an error occurs.",
example:"#DIV/0! → Error"
},

{
title:"IFNA",
keywords:[
"ifna","na error","lookup error","xlookup error","vlookup error",
"not available","replace na","handle na","ifna function","na formula",
"missing value","lookup missing","error handling","na check","if not available"
],
formula:'=IFNA(XLOOKUP(A2,D:D,E:E),"Not Found")',
explanation:"Handles only #N/A errors.",
example:"Unknown ID → Not Found"
},

{
title:"AND",
keywords:[
"and","multiple conditions","all true","logical and","both conditions",
"condition check","logical function","excel and","true false","all criteria",
"and formula","multiple logical test","condition formula","and function","logical test"
],
formula:'=AND(A2>=50,B2="Present")',
explanation:"Returns TRUE if all conditions are TRUE.",
example:"TRUE"
},

{
title:"OR",
keywords:[
"or","any condition","logical or","either condition","one condition",
"excel or","logical function","condition check","true false","or formula",
"multiple conditions","logical test","or function","criteria","decision"
],
formula:'=OR(A2>=50,B2="Present")',
explanation:"Returns TRUE if any condition is TRUE.",
example:"TRUE"
},

{
title:"NOT",
keywords:[
"not","reverse condition","logical not","invert true false","opposite result",
"not formula","excel not","logical formula","condition reverse","not function"
],
formula:'=NOT(A2>50)',
explanation:"Reverses TRUE/FALSE.",
example:"40 → TRUE"
},

{
title:"XOR",
keywords:[
"xor","exclusive or","one true","logical xor","either not both",
"xor function","excel xor","logical formula","condition formula","true false"
],
formula:'=XOR(A2>50,B2>50)',
explanation:"Returns TRUE when only one condition is TRUE.",
example:"TRUE"
},

{
title:"SWITCH",
keywords:[
"switch","multiple options","replace nested if","case statement","switch function",
"excel switch","multiple values","logical switch","text mapping","condition mapping",
"switch formula","value mapping","option selection","switch case","replace if"
],
formula:'=SWITCH(A2,"HR","Human Resource","IT","Information Technology","Others")',
explanation:"Returns a value based on matching cases.",
example:"HR → Human Resource"
},

{
title:"CHOOSE",
keywords:[
"choose","choose value","index selection","pick option","return option",
"choice","choose formula","excel choose","option list","select value",
"index function","multiple options","choose function","selection","choice formula"
],
formula:'=CHOOSE(A2,"Poor","Average","Good","Excellent")',
explanation:"Returns a value based on index number.",
example:"3 → Good"
},

{
title:"COUNTIF",
keywords:[
"countif","count","count condition","count values","conditional count",
"count records","count employee","count formula","excel countif","criteria count",
"count matching","count present","count cells","count data","count function"
],
formula:'=COUNTIF(A:A,"Present")',
explanation:"Counts cells matching one condition.",
example:"150 Present Employees"
},

{
title:"COUNTIFS",
keywords:[
"countifs","multiple count","count multiple conditions","count criteria",
"conditional count","count department","count employee","excel countifs",
"count formula","multiple criteria","count matching","count rows","count data","count values","count function"
],
formula:'=COUNTIFS(A:A,"HR",B:B,"Present")',
explanation:"Counts cells matching multiple conditions.",
example:"HR Present Count"
},

{
title:"SUMIF",
keywords:[
"sumif","conditional sum","sum condition","sum values","sum matching",
"sum employee","sum sales","sum formula","excel sumif","criteria sum",
"add values","sum function","calculate total","sum by name","filtered sum"
],
formula:'=SUMIF(A:A,"Rahul",C:C)',
explanation:"Adds values matching one condition.",
example:"Rahul Sales"
},

{
title:"SUMIFS",
keywords:[
"sumifs","multiple condition sum","sum multiple criteria","conditional total",
"sum values","sum department","sum employee","sum region","excel sumifs",
"sum formula","add values","criteria total","filtered sum","multiple conditions","sum function"
],
formula:'=SUMIFS(C:C,A:A,"Rahul",B:B,"North")',
explanation:"Adds values matching multiple conditions.",
example:"North Region Sales"
},

{
title:"AVERAGEIF",
keywords:[
"averageif","average condition","conditional average","average values","average employee",
"average department","average formula","excel averageif","criteria average",
"average salary","mean condition","filtered average","average function","calculate average","average by department"
],
formula:'=AVERAGEIF(A:A,"HR",B:B)',
explanation:"Returns average matching one condition.",
example:"Average HR Salary"
},

{
title:"AVERAGEIFS",
keywords:[
"averageifs","average multiple conditions","conditional average","multiple criteria average",
"average department","average employee","average salary","excel averageifs",
"average function","filtered average","criteria average","mean values","average data","average formula","average by gender"
],
formula:'=AVERAGEIFS(C:C,A:A,"HR",B:B,"Male")',
explanation:"Returns average matching multiple conditions.",
example:"Average Male HR Salary"
},

{
title:"MAXIFS",
keywords:[
"maxifs","maximum condition","highest value","highest salary","maximum employee",
"criteria max","maximum formula","excel maxifs","largest value","highest department",
"maximum matching","max function","highest sales","filtered max","maximum data"
],
formula:'=MAXIFS(C:C,A:A,"HR")',
explanation:"Returns the maximum matching value.",
example:"Highest HR Salary"
},

{
title:"MINIFS",
keywords:[
"minifs","minimum condition","lowest value","lowest salary","minimum employee",
"criteria minimum","minimum formula","excel minifs","smallest value","lowest department",
"minimum matching","min function","lowest sales","filtered minimum","minimum data"
],
formula:'=MINIFS(C:C,A:A,"HR")',
explanation:"Returns the minimum matching value.",
example:"Lowest HR Salary"
},

{
title:"COUNTBLANK",
keywords:[
"countblank","blank cells","empty cells","count empty","missing values",
"blank formula","excel countblank","count null","empty records","count blanks"
],
formula:'=COUNTBLANK(A:A)',
explanation:"Counts empty cells.",
example:"25 Blank Cells"
},

{
title:"COUNTA",
keywords:[
"counta","non blank","filled cells","count data","count values",
"count non empty","excel counta","filled records","count entries","count formula"
],
formula:'=COUNTA(A:A)',
explanation:"Counts non-empty cells.",
example:"580 Filled Cells"
},

{
title:"ISBLANK",
keywords:[
"isblank","blank","empty","cell empty","blank check",
"empty cell","excel isblank","blank formula","null value","missing cell"
],
formula:'=ISBLANK(A2)',
explanation:"Checks whether a cell is blank.",
example:"TRUE"
},

{
title:"ISNUMBER",
keywords:[
"isnumber","number check","numeric value","is numeric","check number",
"excel isnumber","number formula","numeric formula","validate number","number test"
],
formula:'=ISNUMBER(A2)',
explanation:"Checks whether a value is numeric.",
example:"TRUE"
},

{
title:"ISTEXT",
keywords:[
"istext","text check","check text","text value","string check",
"excel istext","text formula","validate text","string formula","text test"
],
formula:'=ISTEXT(A2)',
explanation:"Checks whether a value is text.",
example:"TRUE"
},

{
title:"ISLOGICAL",
keywords:[
"islogical","logical value","true false","boolean","logical check",
"excel islogical","logical formula","boolean value","logical function","check logical"
],
formula:'=ISLOGICAL(A2)',
explanation:"Checks whether a value is TRUE or FALSE.",
example:"TRUE"
},

{
title:"ISERROR",
keywords:[
"iserror","error check","formula error","excel error","check error",
"is error","error formula","validate error","error value","error function"
],
formula:'=ISERROR(A2)',
explanation:"Checks whether a cell contains an error.",
example:"#VALUE! → TRUE"
},

{
title:"SUM",
keywords:[
"sum","total","add","addition","sum values",
"calculate total","grand total","total sales","add numbers","sum formula",
"excel sum","total amount","calculate sum","sum function","add values"
],
formula:'=SUM(A2:A100)',
explanation:"Adds all numbers in a range.",
example:"Total Sales"
},

{
title:"AVERAGE",
keywords:[
"average","avg","mean","average value","calculate average",
"average salary","average marks","mean value","average formula","excel average",
"average function","calculate mean","average sales","find average","avg formula"
],
formula:'=AVERAGE(A2:A100)',
explanation:"Calculates the average value.",
example:"Average Salary"
},

{
title:"MAX",
keywords:[
"max","maximum","highest","largest","highest value",
"maximum value","top value","highest salary","highest sales","max formula",
"excel max","largest number","maximum function","top score","highest amount"
],
formula:'=MAX(A2:A100)',
explanation:"Returns the highest value.",
example:"Highest Sales"
},

{
title:"MIN",
keywords:[
"min","minimum","lowest","smallest","lowest value",
"minimum value","lowest salary","minimum formula","smallest number","excel min",
"minimum function","lowest score","bottom value","lowest amount","minimum sales"
],
formula:'=MIN(A2:A100)',
explanation:"Returns the lowest value.",
example:"Lowest Salary"
},

{
title:"LARGE",
keywords:[
"large","largest","second highest","third highest","top values",
"nth largest","highest rank","largest formula","excel large","largest number",
"highest salary","top salary","top sales","large function","largest value"
],
formula:'=LARGE(A2:A100,2)',
explanation:"Returns the nth largest value.",
example:"2nd Highest Salary"
},

{
title:"SMALL",
keywords:[
"small","smallest","lowest value","third lowest","nth smallest",
"smallest number","bottom values","lowest rank","small formula","excel small",
"small function","lowest salary","lowest marks","bottom score","smallest value"
],
formula:'=SMALL(A2:A100,3)',
explanation:"Returns the nth smallest value.",
example:"3rd Lowest Score"
},

{
title:"RANK",
keywords:[
"rank","ranking","position","employee rank","sales rank",
"rank formula","excel rank","position number","ranking function","order",
"highest position","lowest position","rank value","leaderboard","score rank"
],
formula:'=RANK(A2,$A$2:$A$100)',
explanation:"Returns the rank of a value.",
example:"Sales Rank"
},

{
title:"RANK.EQ",
keywords:[
"rank eq","rank equal","duplicate rank","equal rank","ranking",
"excel rank eq","rank formula","employee rank","sales rank","rank function",
"same rank","equal position","duplicate values","rank eq formula","ranking values"
],
formula:'=RANK.EQ(A2,$A$2:$A$100)',
explanation:"Returns equal ranking.",
example:"Rank Position"
},

{
title:"RANK.AVG",
keywords:[
"rank avg","average rank","duplicate ranking","average position","rank duplicates",
"excel rank avg","ranking formula","rank average","equal values","ranking function",
"average ranking","duplicate scores","rank formula","average position","rank calculation"
],
formula:'=RANK.AVG(A2,$A$2:$A$100)',
explanation:"Returns average rank for duplicates.",
example:"Average Rank"
},

{
title:"ROUND",
keywords:[
"round","round off","decimal","2 decimals","nearest value",
"round formula","excel round","round number","decimal places","round function",
"rounding","format decimals","number format","rounded value","decimal rounding"
],
formula:'=ROUND(A2,2)',
explanation:"Rounds a number to specified decimal places.",
example:"125.456 → 125.46"
},

{
title:"ROUNDUP",
keywords:[
"roundup","round up","always round up","increase value","next integer",
"excel roundup","round formula","rounding","ceiling","higher number",
"decimal rounding","math rounding","round upward","round function","whole number"
],
formula:'=ROUNDUP(A2,0)',
explanation:"Rounds a number upward.",
example:"12.2 → 13"
},

{
title:"ROUNDDOWN",
keywords:[
"rounddown","round down","always round down","lower value","previous integer",
"excel rounddown","round formula","rounding","floor","smaller number",
"decimal rounding","math rounding","round downward","round function","whole number"
],
formula:'=ROUNDDOWN(A2,0)',
explanation:"Rounds a number downward.",
example:"12.9 → 12"
},

{
title:"CEILING",
keywords:[
"ceiling","nearest multiple","round multiple","round up multiple","multiple of 5",
"excel ceiling","higher multiple","ceiling function","rounding","math",
"next multiple","ceiling formula","nearest higher","multiple rounding","number rounding"
],
formula:'=CEILING(A2,5)',
explanation:"Rounds up to the nearest multiple.",
example:"22 → 25"
},

{
title:"FLOOR",
keywords:[
"floor","nearest lower multiple","round down multiple","multiple of 5",
"excel floor","lower multiple","floor function","math","previous multiple",
"floor formula","number rounding","rounding","nearest lower","multiple rounding","lowest multiple"
],
formula:'=FLOOR(A2,5)',
explanation:"Rounds down to the nearest multiple.",
example:"22 → 20"
},

{
title:"MOD",
keywords:[
"mod","remainder","division remainder","odd even","modulus",
"math","excel mod","balance","divide remainder","mod function",
"remainder formula","number division","check odd even","division formula","mod calculation"
],
formula:'=MOD(A2,2)',
explanation:"Returns the remainder after division.",
example:"15 → 1"
},

{
title:"QUOTIENT",
keywords:[
"quotient","integer division","whole division","divide numbers","quotient formula",
"excel quotient","division result","whole number division","math","divide function",
"integer result","remove remainder","quotient function","division calculation","number division"
],
formula:'=QUOTIENT(A2,B2)',
explanation:"Returns the integer portion of a division.",
example:"25/4 → 6"
},

{
title:"ABS",
keywords:[
"abs","absolute","positive value","remove negative","absolute value",
"excel abs","negative to positive","math","absolute function","positive number",
"minus to plus","abs formula","number conversion","absolute number","convert positive"
],
formula:'=ABS(A2)',
explanation:"Returns the absolute value.",
example:"-15 → 15"
},

{
title:"SQRT",
keywords:[
"sqrt","square root","root","math","square root formula",
"excel sqrt","root calculation","sqrt function","calculate root","number root",
"root value","square","math function","find root","square root value"
],
formula:'=SQRT(A2)',
explanation:"Returns the square root.",
example:"144 → 12"
},

{
title:"POWER",
keywords:[
"power","exponent","square","cube","raise power",
"excel power","power formula","power function","calculate power","exponential",
"number power","square value","cube value","math formula","power calculation"
],
formula:'=POWER(A2,2)',
explanation:"Raises a number to a power.",
example:"5² → 25"
},

{
title:"PRODUCT",
keywords:[
"product","multiply","multiplication","multiply values","times",
"excel product","product formula","multiply cells","math","product function",
"multiplication formula","calculate product","multiply numbers","total multiplication","number multiplication"
],
formula:'=PRODUCT(A2:A5)',
explanation:"Multiplies all numbers in a range.",
example:"2×3×4×5"
},

{
title:"SUBTOTAL",
keywords:[
"subtotal","filtered total","subtotal formula","filtered sum","visible rows",
"excel subtotal","subtotal function","sum visible","filter total","subtotal calculation",
"ignore hidden rows","total after filter","subtotal values","filter calculation","subtotal data"
],
formula:'=SUBTOTAL(9,A2:A100)',
explanation:"Calculates subtotal ignoring filtered rows.",
example:"Filtered Total"
},

{
title:"AGGREGATE",
keywords:[
"aggregate","aggregate formula","advanced subtotal","ignore errors","filtered sum",
"excel aggregate","aggregate function","visible rows","advanced calculation","summary",
"aggregate values","filter calculation","error handling","aggregate total","aggregation"
],
formula:'=AGGREGATE(9,5,A2:A100)',
explanation:"Performs advanced aggregation.",
example:"Filtered Sum"
},

{
title:"RAND",
keywords:[
"rand","random","random decimal","random number","generate random",
"excel rand","random formula","decimal random","random value","number generator",
"random function","generate decimal","random data","random value generator","randomize"
],
formula:'=RAND()',
explanation:"Returns a random decimal number.",
example:"0.7345"
},

{
title:"RANDBETWEEN",
keywords:[
"randbetween","random integer","random whole number","random value","generate random",
"excel randbetween","random formula","number generator","random between","integer random",
"random function","generate integer","random numbers","random range","random data"
],
formula:'=RANDBETWEEN(1,100)',
explanation:"Returns a random integer.",
example:"58"
},

{
title:"INT",
keywords:[
"int","integer","remove decimal","whole number","round down",
"excel int","integer formula","math","truncate","remove fraction",
"int function","convert integer","whole value","integer conversion","number truncation"
],
formula:'=INT(A2)',
explanation:"Rounds a number down to the nearest integer.",
example:"12.95 → 12"
},

{
title:"TODAY",
keywords:[
"today","current date","today date","present date","system date",
"current day","today function","excel today","today formula","date today",
"current date formula","today value","today excel","show today's date","live date"
],
formula:'=TODAY()',
explanation:"Returns today's date.",
example:"21-Jul-2026"
},

{
title:"NOW",
keywords:[
"now","current time","current date time","system time","live time",
"today time","current datetime","excel now","now formula","date and time",
"current timestamp","live date time","time now","current clock","today now"
],
formula:'=NOW()',
explanation:"Returns the current date and time.",
example:"21-Jul-2026 09:30 AM"
},

{
title:"DATE",
keywords:[
"date","create date","date formula","make date","date value",
"excel date","year month day","date function","generate date","calendar date",
"date creation","custom date","valid date","date format","excel calendar"
],
formula:'=DATE(2026,7,21)',
explanation:"Creates a valid Excel date.",
example:"21-Jul-2026"
},

{
title:"TIME",
keywords:[
"time","create time","time formula","time value","hour minute second",
"excel time","generate time","clock time","time function","custom time",
"make time","time format","time calculation","excel clock","time entry"
],
formula:'=TIME(9,30,0)',
explanation:"Creates a valid Excel time.",
example:"09:30 AM"
},

{
title:"DAY",
keywords:[
"day","day number","day of month","extract day","date day",
"excel day","day formula","calendar day","date function","day value"
],
formula:'=DAY(A2)',
explanation:"Returns the day of the month.",
example:"21"
},

{
title:"MONTH",
keywords:[
"month","month number","extract month","date month","calendar month",
"excel month","month formula","month value","date function","financial month"
],
formula:'=MONTH(A2)',
explanation:"Returns the month number.",
example:"7"
},

{
title:"YEAR",
keywords:[
"year","extract year","date year","calendar year","financial year",
"excel year","year formula","year value","date function","current year"
],
formula:'=YEAR(A2)',
explanation:"Returns the year.",
example:"2026"
},

{
title:"HOUR",
keywords:[
"hour","extract hour","time hour","current hour","hour number",
"excel hour","hour formula","clock hour","time function","hour value"
],
formula:'=HOUR(A2)',
explanation:"Returns the hour.",
example:"9"
},

{
title:"MINUTE",
keywords:[
"minute","extract minute","time minute","current minute","minute value",
"excel minute","minute formula","clock minute","time function","minutes"
],
formula:'=MINUTE(A2)',
explanation:"Returns the minute.",
example:"30"
},

{
title:"SECOND",
keywords:[
"second","extract second","time second","current second","seconds",
"excel second","second formula","clock second","time function","second value"
],
formula:'=SECOND(A2)',
explanation:"Returns the second.",
example:"45"
},

{
title:"WEEKDAY",
keywords:[
"weekday","day name","week day","day of week","weekday number",
"excel weekday","monday tuesday","week formula","weekday function","calendar day"
],
formula:'=WEEKDAY(A2)',
explanation:"Returns the weekday number.",
example:"Tuesday → 3"
},

{
title:"WEEKNUM",
keywords:[
"weeknum","week number","week of year","calendar week","week calculation",
"excel weeknum","week formula","current week","week value","week function"
],
formula:'=WEEKNUM(A2)',
explanation:"Returns the week number of the year.",
example:"Week 30"
},

{
title:"ISOWEEKNUM",
keywords:[
"isoweeknum","iso week","iso week number","week standard","iso calendar",
"excel isoweeknum","week formula","international week","iso week function","week value"
],
formula:'=ISOWEEKNUM(A2)',
explanation:"Returns the ISO week number.",
example:"Week 30"
},

{
title:"EDATE",
keywords:[
"edate","add months","subtract months","future date","previous month",
"excel edate","month calculation","date after months","month formula","date shift"
],
formula:'=EDATE(A2,3)',
explanation:"Adds or subtracts months.",
example:"21-Oct-2026"
},

{
title:"EOMONTH",
keywords:[
"eomonth","end of month","last day month","month end","month closing",
"excel eomonth","month formula","last date","month calculation","month finish"
],
formula:'=EOMONTH(A2,0)',
explanation:"Returns the last day of the month.",
example:"31-Jul-2026"
},

{
title:"WORKDAY",
keywords:[
"workday","working day","business day","next working day","office day",
"excel workday","holiday calculation","working date","business date","workday formula"
],
formula:'=WORKDAY(A2,10)',
explanation:"Returns a working day after a given number of days.",
example:"10 Business Days Later"
},

{
title:"NETWORKDAYS",
keywords:[
"networkdays","working days","business days","office days","working day count",
"excel networkdays","holiday calculation","days between","working period","business calculation"
],
formula:'=NETWORKDAYS(A2,B2)',
explanation:"Counts working days between two dates.",
example:"22 Working Days"
},

{
title:"DAYS",
keywords:[
"days","difference in days","date difference","days between","count days",
"excel days","day calculation","date duration","number of days","days formula"
],
formula:'=DAYS(B2,A2)',
explanation:"Returns the number of days between two dates.",
example:"45 Days"
},

{
title:"DATEDIF",
keywords:[
"datedif","date difference","age calculation","years months days","service period",
"excel datedif","experience calculation","difference between dates","date duration","age formula"
],
formula:'=DATEDIF(A2,B2,"Y")',
explanation:"Calculates the difference between two dates.",
example:"5 Years"
},

{
title:"YEARFRAC",
keywords:[
"yearfrac","fraction of year","year difference","financial year","interest calculation",
"excel yearfrac","year formula","decimal year","year duration","financial formula"
],
formula:'=YEARFRAC(A2,B2)',
explanation:"Returns the fraction of a year.",
example:"2.75 Years"
},

{
title:"DATEVALUE",
keywords:[
"datevalue","text to date","convert date","date conversion","date text",
"excel datevalue","date formula","convert text date","valid date","date parsing"
],
formula:'=DATEVALUE("21-Jul-2026")',
explanation:"Converts text into a valid date.",
example:"21-Jul-2026"
},

{
title:"TIMEVALUE",
keywords:[
"timevalue","text to time","convert time","time conversion","time text",
"excel timevalue","time formula","convert text time","valid time","time parsing"
],
formula:'=TIMEVALUE("09:30 AM")',
explanation:"Converts text into a valid time.",
example:"09:30 AM"
},

{
title:"MONTHNAME",
keywords:[
"month name","month text","display month","month from date","month word",
"mmm","mmmm","text month","excel month name","month display"
],
formula:'=TEXT(A2,"mmmm")',
explanation:"Returns the full month name.",
example:"July"
},

{
title:"DAYNAME",
keywords:[
"day name","weekday name","day text","display day","week day name",
"day word","excel day name","weekday text","day display","weekday full name"
],
formula:'=TEXT(A2,"dddd")',
explanation:"Returns the full weekday name.",
example:"Monday"
},

{
title:"CURRENT YEAR",
keywords:[
"current year","this year","present year","today year","year today",
"current financial year","system year","year now","excel current year","year formula"
],
formula:'=YEAR(TODAY())',
explanation:"Returns the current year.",
example:"2026"
},

{
title:"PMT",
keywords:[
"pmt","emi","loan emi","monthly payment","loan payment",
"emi calculator","mortgage payment","installment","payment formula","loan installment"
],
formula:'=PMT(10%/12,60,-500000)',
explanation:"Calculates the monthly loan payment.",
example:"₹500000 Loan → EMI"
},

{
title:"PV",
keywords:[
"pv","present value","loan value","current value","discount value",
"present amount","financial value","investment today","present worth","finance pv"
],
formula:'=PV(10%,5,-10000)',
explanation:"Returns the present value of an investment.",
example:"Current Investment Value"
},

{
title:"FV",
keywords:[
"fv","future value","investment growth","future amount","future investment",
"investment value","future worth","compound value","finance fv","future money"
],
formula:'=FV(10%,5,-10000)',
explanation:"Returns the future value of an investment.",
example:"Investment after 5 Years"
},

{
title:"RATE",
keywords:[
"rate","interest rate","loan interest","investment rate","calculate rate",
"finance rate","interest formula","annual rate","monthly rate","rate function"
],
formula:'=RATE(60,-10000,500000)',
explanation:"Calculates the interest rate.",
example:"Loan Interest Rate"
},

{
title:"NPER",
keywords:[
"nper","loan period","number of payments","payment period","loan tenure",
"investment duration","finance nper","payment count","emi months","loan months"
],
formula:'=NPER(10%/12,-10000,500000)',
explanation:"Returns the number of payment periods.",
example:"60 Months"
},

{
title:"IPMT",
keywords:[
"ipmt","interest payment","loan interest amount","interest portion","emi interest",
"finance ipmt","monthly interest","payment interest","interest calculation","loan interest"
],
formula:'=IPMT(10%/12,1,60,-500000)',
explanation:"Returns the interest payment for a period.",
example:"Month 1 Interest"
},

{
title:"PPMT",
keywords:[
"ppmt","principal payment","loan principal","principal portion","emi principal",
"finance ppmt","monthly principal","loan balance","principal calculation","payment principal"
],
formula:'=PPMT(10%/12,1,60,-500000)',
explanation:"Returns the principal payment for a period.",
example:"Month 1 Principal"
},

{
title:"NPV",
keywords:[
"npv","net present value","cash flow","investment analysis","project value",
"discounted cash flow","finance npv","investment return","npv formula","capital budgeting"
],
formula:'=NPV(10%,B2:B6)',
explanation:"Calculates the net present value.",
example:"Project Evaluation"
},

{
title:"IRR",
keywords:[
"irr","internal rate of return","investment return","project return","finance irr",
"cash flow return","irr formula","profitability","investment analysis","irr calculation"
],
formula:'=IRR(B2:B10)',
explanation:"Calculates the internal rate of return.",
example:"Investment Return %"
},

{
title:"XIRR",
keywords:[
"xirr","irregular cash flow","investment xirr","actual dates","cash flow return",
"finance xirr","annual return","investment performance","xirr formula","cashflow"
],
formula:'=XIRR(B2:B10,A2:A10)',
explanation:"Calculates IRR for irregular cash flows.",
example:"Investment Analysis"
},

{
title:"XNPV",
keywords:[
"xnpv","net present value dates","discounted cash flow","irregular cash flow",
"finance xnpv","investment valuation","cash flow dates","xnpv formula","project value","valuation"
],
formula:'=XNPV(10%,B2:B10,A2:A10)',
explanation:"Calculates NPV using actual dates.",
example:"Investment Valuation"
},

{
title:"SLN",
keywords:[
"sln","straight line depreciation","asset depreciation","depreciation","asset value",
"finance sln","book value","asset life","depreciation formula","fixed asset"
],
formula:'=SLN(100000,10000,5)',
explanation:"Calculates straight-line depreciation.",
example:"Annual Depreciation"
},

{
title:"DB",
keywords:[
"db","declining balance","depreciation db","asset depreciation","finance db",
"book value","fixed asset","declining depreciation","asset cost","depreciation function"
],
formula:'=DB(100000,10000,5,1)',
explanation:"Calculates declining balance depreciation.",
example:"Year 1 Depreciation"
},

{
title:"DDB",
keywords:[
"ddb","double declining balance","accelerated depreciation","asset depreciation",
"finance ddb","fixed asset","depreciation","asset cost","ddb formula","book value"
],
formula:'=DDB(100000,10000,5,1)',
explanation:"Calculates double declining balance depreciation.",
example:"Accelerated Depreciation"
},

{
title:"SYD",
keywords:[
"syd","sum of years digits","asset depreciation","depreciation method",
"finance syd","asset value","book value","depreciation formula","fixed asset","asset life"
],
formula:'=SYD(100000,10000,5,1)',
explanation:"Calculates depreciation using the Sum-of-Years' Digits method.",
example:"Year 1 Depreciation"
},

{
title:"EFFECT",
keywords:[
"effect","effective interest","effective annual rate","ear","annual interest",
"finance effect","interest calculation","effective rate","bank interest","compound rate"
],
formula:'=EFFECT(10%,12)',
explanation:"Returns the effective annual interest rate.",
example:"10.47%"
},

{
title:"NOMINAL",
keywords:[
"nominal","nominal interest","annual rate","nominal rate","finance nominal",
"interest rate","bank interest","loan rate","nominal formula","interest calculation"
],
formula:'=NOMINAL(10.47%,12)',
explanation:"Returns the nominal annual interest rate.",
example:"10%"
},

{
title:"CUMIPMT",
keywords:[
"cumipmt","cumulative interest","loan interest total","interest paid","loan calculation",
"finance cumipmt","emi interest","interest summary","loan interest","payment interest"
],
formula:'=CUMIPMT(10%/12,60,500000,1,12,0)',
explanation:"Calculates cumulative interest paid.",
example:"Interest Paid in First Year"
},

{
title:"CUMPRINC",
keywords:[
"cumprinc","cumulative principal","principal paid","loan principal total","emi principal",
"finance cumprinc","loan calculation","principal summary","loan payment","principal amount"
],
formula:'=CUMPRINC(10%/12,60,500000,1,12,0)',
explanation:"Calculates cumulative principal paid.",
example:"Principal Paid in First Year"
},

{
title:"PRICE",
keywords:[
"price","bond price","security price","bond valuation","finance price",
"investment price","bond formula","coupon bond","bond value","price function"
],
formula:'=PRICE(A2,B2,5%,6%,100,2,0)',
explanation:"Returns the price of a bond.",
example:"Bond Price"
},

{
title:"YIELD",
keywords:[
"yield","bond yield","investment yield","security yield","finance yield",
"bond return","yield formula","coupon yield","investment return","bond calculation"
],
formula:'=YIELD(A2,B2,5%,95,100,2,0)',
explanation:"Returns the annual bond yield.",
example:"Bond Yield %"
},

{
title:"DISC",
keywords:[
"disc","discount rate","bond discount","security discount","discount calculation",
"finance disc","treasury bill","investment discount","discount formula","bond finance"
],
formula:'=DISC(A2,B2,97,100,0)',
explanation:"Returns the discount rate of a security.",
example:"Treasury Bill Discount"
},

{
title:"DOLLARDE",
keywords:[
"dollarde","fraction to decimal","bond quotation","price conversion","financial conversion",
"finance dollarde","decimal conversion","bond price","fraction value","decimal value"
],
formula:'=DOLLARDE(100.25,16)',
explanation:"Converts a fractional dollar value to decimal.",
example:"100.15625"
},

{
title:"DOLLARFR",
keywords:[
"dollarfr","decimal to fraction","bond quotation","price fraction","financial conversion",
"finance dollarfr","fraction conversion","bond price","fraction format","price display"
],
formula:'=DOLLARFR(100.15625,16)',
explanation:"Converts a decimal dollar value to fractional.",
example:"100.25"
},

{
title:"ACCRINT",
keywords:[
"accrint","accrued interest","bond interest","security interest","interest accrued",
"finance accrint","coupon interest","bond calculation","investment interest","interest formula"
],
formula:'=ACCRINT(A2,B2,C2,5%,100,2)',
explanation:"Returns accrued interest for a security.",
example:"Accrued Bond Interest"
},

{
title:"MEDIAN",
keywords:[
"median","middle value","median formula","find median","average middle",
"excel median","middle number","statistics median","median function","median value"
],
formula:'=MEDIAN(A2:A100)',
explanation:"Returns the middle value in a dataset.",
example:"Median Salary"
},

{
title:"MODE.SNGL",
keywords:[
"mode","mode single","most repeated","most frequent","common value",
"excel mode","mode formula","mode function","highest frequency","repeated value"
],
formula:'=MODE.SNGL(A2:A100)',
explanation:"Returns the most frequently occurring value.",
example:"Most Common Salary"
},

{
title:"MODE.MULT",
keywords:[
"mode multiple","multiple modes","duplicate values","most frequent values",
"excel mode mult","mode mult","mode function","statistics","common values","repeated numbers"
],
formula:'=MODE.MULT(A2:A100)',
explanation:"Returns multiple most frequent values.",
example:"Repeated Scores"
},

{
title:"STDEV.S",
keywords:[
"standard deviation","stdev","sample deviation","sample standard deviation",
"excel stdev","stdev sample","statistics","spread","variation","deviation"
],
formula:'=STDEV.S(A2:A100)',
explanation:"Calculates sample standard deviation.",
example:"Salary Variation"
},

{
title:"STDEV.P",
keywords:[
"population deviation","stdev population","standard deviation population",
"excel stdev.p","population statistics","deviation","spread","population variance","statistics","stdevp"
],
formula:'=STDEV.P(A2:A100)',
explanation:"Calculates population standard deviation.",
example:"Population Analysis"
},

{
title:"VAR.S",
keywords:[
"variance sample","sample variance","var sample","variance",
"excel var.s","statistics","sample spread","variance formula","variation","sample"
],
formula:'=VAR.S(A2:A100)',
explanation:"Returns sample variance.",
example:"Salary Variance"
},

{
title:"VAR.P",
keywords:[
"variance population","population variance","var population","excel var.p",
"population statistics","variance","spread","variation","population","statistics"
],
formula:'=VAR.P(A2:A100)',
explanation:"Returns population variance.",
example:"Population Variance"
},

{
title:"PERCENTILE",
keywords:[
"percentile","90 percentile","95 percentile","percentile rank",
"excel percentile","statistics","top percentile","percentile formula","distribution","ranking"
],
formula:'=PERCENTILE.INC(A2:A100,0.9)',
explanation:"Returns the specified percentile.",
example:"90th Percentile Salary"
},

{
title:"QUARTILE",
keywords:[
"quartile","first quartile","third quartile","quartiles","q1","q2","q3",
"excel quartile","statistics","distribution"
],
formula:'=QUARTILE.INC(A2:A100,1)',
explanation:"Returns the quartile of a dataset.",
example:"Q1 Salary"
},

{
title:"FREQUENCY",
keywords:[
"frequency","frequency distribution","count frequency","histogram",
"excel frequency","frequency formula","distribution","statistics","count groups","bins"
],
formula:'=FREQUENCY(A2:A100,C2:C10)',
explanation:"Calculates how often values occur.",
example:"Marks Distribution"
},

{
title:"CORREL",
keywords:[
"correlation","correl","relationship","correlation coefficient",
"excel correl","statistics","compare datasets","correlation analysis","linear relationship","corr"
],
formula:'=CORREL(A2:A100,B2:B100)',
explanation:"Returns the correlation coefficient.",
example:"Sales vs Profit"
},

{
title:"COVARIANCE.S",
keywords:[
"covariance","covariance sample","sample covariance","relationship",
"excel covariance","statistics","covariance s","dataset comparison","covariance formula","sample relation"
],
formula:'=COVARIANCE.S(A2:A100,B2:B100)',
explanation:"Returns sample covariance.",
example:"Sales & Profit"
},

{
title:"COVARIANCE.P",
keywords:[
"covariance population","population covariance","covariance p",
"excel covariance.p","statistics","relationship","population analysis","covariance formula","dataset"
],
formula:'=COVARIANCE.P(A2:A100,B2:B100)',
explanation:"Returns population covariance.",
example:"Population Relationship"
},

{
title:"SKEW",
keywords:[
"skew","data skewness","distribution skew","statistics skew",
"excel skew","asymmetry","distribution","skew formula","statistical analysis","data distribution"
],
formula:'=SKEW(A2:A100)',
explanation:"Measures the asymmetry of data.",
example:"Sales Distribution"
},

{
title:"KURT",
keywords:[
"kurtosis","kurt","distribution peak","statistics kurtosis",
"excel kurt","distribution analysis","peakedness","kurt formula","statistical function","data shape"
],
formula:'=KURT(A2:A100)',
explanation:"Returns the kurtosis of a dataset.",
example:"Distribution Shape"
},

{
title:"GEOMEAN",
keywords:[
"geometric mean","geomean","growth average","compound average",
"excel geomean","statistics","investment growth","geometric average","mean","growth rate"
],
formula:'=GEOMEAN(A2:A100)',
explanation:"Returns the geometric mean.",
example:"Average Growth Rate"
},

{
title:"HARMEAN",
keywords:[
"harmonic mean","harmean","average rate","speed average",
"excel harmean","statistics","harmonic average","mean","rate calculation","average"
],
formula:'=HARMEAN(A2:A100)',
explanation:"Returns the harmonic mean.",
example:"Average Speed"
},

{
title:"TRIMMEAN",
keywords:[
"trimmean","trim mean","remove outliers","trim average",
"excel trimmean","statistics","average without outliers","mean","clean average","outlier"
],
formula:'=TRIMMEAN(A2:A100,0.2)',
explanation:"Returns the mean excluding outliers.",
example:"Adjusted Average"
},

{
title:"NORM.DIST",
keywords:[
"normal distribution","norm dist","normal probability","bell curve",
"excel norm.dist","statistics","distribution","probability","normal curve","norm"
],
formula:'=NORM.DIST(A2,50,10,TRUE)',
explanation:"Returns the normal distribution.",
example:"Probability Value"
},

{
title:"NORM.INV",
keywords:[
"normal inverse","norm inv","inverse normal","probability inverse",
"excel norm.inv","statistics","distribution","inverse probability","normal calculation","norminv"
],
formula:'=NORM.INV(0.95,50,10)',
explanation:"Returns the inverse normal distribution.",
example:"95% Confidence Value"
},

{
title:"PERCENTRANK",
keywords:[
"percent rank","ranking percentage","percentile rank","rank percentage",
"excel percentrank","statistics","ranking","percent position","rank function","distribution"
],
formula:'=PERCENTRANK.INC(A2:A100,A2)',
explanation:"Returns the percentage rank.",
example:"Top 10%"
},

{
title:"Z.TEST",
keywords:[
"z test","ztest","hypothesis testing","z statistic",
"excel ztest","statistics","significance test","probability","sample testing","z value"
],
formula:'=Z.TEST(A2:A100,50)',
explanation:"Returns the one-tailed probability value of a Z-test.",
example:"Hypothesis Test"
},

{
title:"T.TEST",
keywords:[
"t test","ttest","student t test","compare means",
"excel t.test","statistics","hypothesis testing","sample comparison","significance","t statistic"
],
formula:'=T.TEST(A2:A100,B2:B100,2,2)',
explanation:"Returns the probability of a Student's t-test.",
example:"Compare Two Samples"
},

{
title:"FORECAST.LINEAR",
keywords:[
"forecast","forecast linear","prediction","predict value",
"excel forecast","future value","linear forecast","forecast formula","prediction formula","trend prediction"
],
formula:'=FORECAST.LINEAR(2027,B2:B10,A2:A10)',
explanation:"Predicts a future value using linear regression.",
example:"Forecast Next Year Sales"
},

{
title:"TREND",
keywords:[
"trend","trend analysis","future trend","linear trend",
"excel trend","prediction","forecast trend","trend formula","regression","sales trend"
],
formula:'=TREND(B2:B10,A2:A10,A11)',
explanation:"Returns values along a linear trend.",
example:"Predicted Sales"
},

{
title:"FILTER",
keywords:[
"filter","filter data","filter rows","show matching records","filter table",
"excel filter","dynamic filter","filter formula","extract records","filtered list"
],
formula:'=FILTER(A2:D100,C2:C100="HR")',
explanation:"Returns only matching records.",
example:"Only HR Employees"
},

{
title:"UNIQUE",
keywords:[
"unique","remove duplicates","distinct values","unique list","distinct",
"excel unique","unique formula","no duplicates","unique records","different values"
],
formula:'=UNIQUE(A2:A100)',
explanation:"Returns unique values from a range.",
example:"Unique Departments"
},

{
title:"SORT",
keywords:[
"sort","sort ascending","arrange data","sort table","sort values",
"excel sort","ascending","descending","order data","sort formula"
],
formula:'=SORT(A2:D100)',
explanation:"Sorts data in ascending order.",
example:"Employee List"
},

{
title:"SORTBY",
keywords:[
"sortby","sort by column","sort salary","custom sort","sort another column",
"excel sortby","descending sort","dynamic sort","sort formula","highest salary"
],
formula:'=SORTBY(A2:D100,D2:D100,-1)',
explanation:"Sorts data by another column.",
example:"Highest Salary First"
},

{
title:"SEQUENCE",
keywords:[
"sequence","generate numbers","auto numbering","serial number","number sequence",
"excel sequence","dynamic numbers","1 to 100","running numbers","sequence formula"
],
formula:'=SEQUENCE(10)',
explanation:"Generates a sequence of numbers.",
example:"1 to 10"
},

{
title:"TRANSPOSE",
keywords:[
"transpose","rows to columns","columns to rows","rotate table","flip data",
"excel transpose","transpose formula","horizontal vertical","table rotation","matrix"
],
formula:'=TRANSPOSE(A1:D5)',
explanation:"Converts rows into columns.",
example:"Horizontal → Vertical"
},

{
title:"TAKE",
keywords:[
"take","first rows","top records","take rows","top values",
"excel take","dynamic array","take columns","extract top","take formula"
],
formula:'=TAKE(A2:D100,10)',
explanation:"Returns the first rows from an array.",
example:"Top 10 Employees"
},

{
title:"DROP",
keywords:[
"drop","remove first rows","skip rows","ignore top rows","drop records",
"excel drop","dynamic array","remove rows","drop formula","skip data"
],
formula:'=DROP(A2:D100,5)',
explanation:"Removes rows from the beginning.",
example:"Skip First 5 Records"
},

{
title:"CHOOSECOLS",
keywords:[
"choose columns","select columns","pick columns","specific columns",
"excel choosecols","column selection","extract columns","dynamic columns","column filter","choosecols"
],
formula:'=CHOOSECOLS(A2:F100,1,3,5)',
explanation:"Returns selected columns.",
example:"Employee ID, Department & Salary"
},

{
title:"CHOOSEROWS",
keywords:[
"choose rows","select rows","pick rows","specific rows",
"excel chooserows","row selection","extract rows","dynamic rows","choose rows formula","chooserows"
],
formula:'=CHOOSEROWS(A2:D100,1,5,10)',
explanation:"Returns selected rows.",
example:"1st, 5th & 10th Records"
},

{
title:"TOCOL",
keywords:[
"tocol","convert column","single column","flatten column","stack column",
"excel tocol","dynamic arrays","flatten data","column conversion","tocol formula"
],
formula:'=TOCOL(A1:D5)',
explanation:"Converts an array into one column.",
example:"Single Column List"
},

{
title:"TOROW",
keywords:[
"torow","convert row","single row","flatten row","stack row",
"excel torow","dynamic arrays","row conversion","torow formula","flatten table"
],
formula:'=TOROW(A1:D5)',
explanation:"Converts an array into one row.",
example:"Single Row List"
},

{
title:"WRAPROWS",
keywords:[
"wrap rows","convert list to rows","wrap data","array rows","reshape rows",
"excel wraprows","dynamic arrays","wrap formula","rows layout","wraprows"
],
formula:'=WRAPROWS(A1:A12,3)',
explanation:"Wraps values into multiple rows.",
example:"3 Values Per Row"
},

{
title:"WRAPCOLS",
keywords:[
"wrap columns","convert list to columns","wrap data","array columns","reshape columns",
"excel wrapcols","dynamic arrays","wrapcols formula","column layout","wrap"
],
formula:'=WRAPCOLS(A1:A12,3)',
explanation:"Wraps values into multiple columns.",
example:"3 Values Per Column"
},

{
title:"VSTACK",
keywords:[
"vstack","append tables","merge tables","combine tables","append rows",
"excel vstack","vertical merge","stack rows","combine datasets","merge data"
],
formula:'=VSTACK(A2:D10,F2:I10)',
explanation:"Stacks arrays vertically.",
example:"Append Two Tables"
},

{
title:"HSTACK",
keywords:[
"hstack","merge columns","combine columns","append columns","horizontal merge",
"excel hstack","stack columns","combine datasets","merge sheets","hstack formula"
],
formula:'=HSTACK(A2:B10,D2:E10)',
explanation:"Stacks arrays horizontally.",
example:"Merge Two Tables"
},

{
title:"EXPAND",
keywords:[
"expand","resize array","increase array","expand formula","dynamic array",
"excel expand","array resize","expand rows","expand columns","expand data"
],
formula:'=EXPAND(A1:B5,10,5)',
explanation:"Expands an array to a specified size.",
example:"10 Rows × 5 Columns"
},

{
title:"TEXTSPLIT",
keywords:[
"textsplit","split text","split delimiter","split columns","split comma",
"excel textsplit","separate text","split values","csv split","text split"
],
formula:'=TEXTSPLIT(A2,",")',
explanation:"Splits text into multiple columns.",
example:"HR,IT,Finance"
},

{
title:"TEXTBEFORE",
keywords:[
"textbefore","before delimiter","left of dash","before comma","extract before",
"excel textbefore","text before","before symbol","extract left","delimiter"
],
formula:'=TEXTBEFORE(A2,"-")',
explanation:"Returns text before a delimiter.",
example:"EMP-1001 → EMP"
},

{
title:"TEXTAFTER",
keywords:[
"textafter","after delimiter","right of dash","after comma","extract after",
"excel textafter","text after","after symbol","extract right","delimiter"
],
formula:'=TEXTAFTER(A2,"-")',
explanation:"Returns text after a delimiter.",
example:"EMP-1001 → 1001"
},

{
title:"TEXTJOIN",
keywords:[
"textjoin","join text","merge text","combine cells","concatenate range",
"excel textjoin","join values","merge values","combine strings","text join"
],
formula:'=TEXTJOIN(", ",TRUE,A2:A10)',
explanation:"Joins multiple values with a delimiter.",
example:"HR, IT, Finance"
},

{
title:"LET",
keywords:[
"let","let function","variables","formula variables","optimize formula",
"excel let","advanced formula","named variable","formula optimization","let formula"
],
formula:'=LET(x,A2:B10,SUM(x))',
explanation:"Assigns variables inside formulas.",
example:"Simplified Complex Formula"
},

{
title:"LAMBDA",
keywords:[
"lambda","custom function","user defined formula","reusable formula",
"excel lambda","lambda function","custom excel function","advanced excel","formula function","lambda excel"
],
formula:'=LAMBDA(x,x*10)(5)',
explanation:"Creates reusable custom functions.",
example:"Returns 50"
},

{
title:"MAP",
keywords:[
"map","apply formula","transform array","array calculation","map function",
"excel map","dynamic arrays","apply lambda","map lambda","array transform"
],
formula:'=MAP(A2:A10,LAMBDA(x,x*2))',
explanation:"Applies a function to every value.",
example:"Double Every Number"
},

{
title:"REDUCE",
keywords:[
"reduce","reduce array","accumulate","running calculation","array reduce",
"excel reduce","lambda reduce","dynamic arrays","sum array","reduce function"
],
formula:'=REDUCE(0,A2:A10,LAMBDA(a,b,a+b))',
explanation:"Reduces an array to one value.",
example:"Running Total"
},

{
title:"DSUM",
keywords:[
"dsum","database sum","sum database","sum with criteria","conditional sum",
"excel dsum","database formula","criteria sum","filter sum","database total"
],
formula:'=DSUM(A1:F100,"Salary",H1:H2)',
explanation:"Adds values that match specified criteria.",
example:"Total HR Salary"
},

{
title:"DAVERAGE",
keywords:[
"daverage","database average","average database","average with criteria","conditional average",
"excel daverage","database formula","criteria average","database mean","filtered average"
],
formula:'=DAVERAGE(A1:F100,"Salary",H1:H2)',
explanation:"Calculates the average of matching records.",
example:"Average IT Salary"
},

{
title:"DCOUNT",
keywords:[
"dcount","database count","count database","count records","count criteria",
"excel dcount","database formula","filtered count","database rows","criteria count"
],
formula:'=DCOUNT(A1:F100,"Salary",H1:H2)',
explanation:"Counts numeric records matching criteria.",
example:"Employees with Salary"
},

{
title:"DCOUNTA",
keywords:[
"dcounta","count non blank","database nonblank","count text database","database values",
"excel dcounta","database formula","non empty records","criteria count","filled records"
],
formula:'=DCOUNTA(A1:F100,"Employee",H1:H2)',
explanation:"Counts non-empty records matching criteria.",
example:"Total HR Employees"
},

{
title:"DGET",
keywords:[
"dget","database lookup","single record","database search","find record",
"excel dget","lookup database","criteria lookup","employee lookup","database value"
],
formula:'=DGET(A1:F100,"Salary",H1:H2)',
explanation:"Returns one record that matches the criteria.",
example:"Rahul Salary"
},

{
title:"DMAX",
keywords:[
"dmax","database max","highest value database","maximum database","database highest",
"excel dmax","max with criteria","database formula","highest salary","criteria max"
],
formula:'=DMAX(A1:F100,"Salary",H1:H2)',
explanation:"Returns the maximum value matching criteria.",
example:"Highest HR Salary"
},

{
title:"DMIN",
keywords:[
"dmin","database min","lowest value database","minimum database","database lowest",
"excel dmin","min with criteria","database formula","lowest salary","criteria minimum"
],
formula:'=DMIN(A1:F100,"Salary",H1:H2)',
explanation:"Returns the minimum value matching criteria.",
example:"Lowest HR Salary"
},

{
title:"DPRODUCT",
keywords:[
"dproduct","database product","multiply database","product criteria","database multiply",
"excel dproduct","criteria product","database formula","multiply values","filtered product"
],
formula:'=DPRODUCT(A1:F100,"Sales",H1:H2)',
explanation:"Multiplies values matching criteria.",
example:"Sales Product"
},

{
title:"DSTDEV",
keywords:[
"dstdev","database standard deviation","sample deviation database","database statistics",
"excel dstdev","criteria deviation","database formula","sample stdev","database spread","variation"
],
formula:'=DSTDEV(A1:F100,"Salary",H1:H2)',
explanation:"Returns the sample standard deviation.",
example:"HR Salary Variation"
},

{
title:"DSTDEVP",
keywords:[
"dstdevp","population deviation database","database stdev population","population statistics",
"excel dstdevp","criteria deviation","database formula","population spread","database analysis","stdev population"
],
formula:'=DSTDEVP(A1:F100,"Salary",H1:H2)',
explanation:"Returns the population standard deviation.",
example:"Company Salary Analysis"
},

{
title:"DVAR",
keywords:[
"dvar","database variance","sample variance database","variance criteria","database statistics",
"excel dvar","database formula","sample variance","variation","database spread"
],
formula:'=DVAR(A1:F100,"Salary",H1:H2)',
explanation:"Returns the sample variance.",
example:"HR Salary Variance"
},

{
title:"DVARP",
keywords:[
"dvarp","population variance database","database variance population","population statistics",
"excel dvarp","database formula","variance","database spread","population analysis","variance function"
],
formula:'=DVARP(A1:F100,"Salary",H1:H2)',
explanation:"Returns the population variance.",
example:"Company Salary Variance"
},

{
title:"COUNTIF",
keywords:[
"countif","count matching","count condition","count criteria","conditional count",
"excel countif","attendance count","status count","count formula","count records"
],
formula:'=COUNTIF(B:B,"Present")',
explanation:"Counts cells that match one condition.",
example:"Present Employees"
},

{
title:"COUNTIFS",
keywords:[
"countifs","multiple conditions","count multiple criteria","countifs formula","conditional counting",
"excel countifs","department attendance","criteria count","multi condition count","count records"
],
formula:'=COUNTIFS(B:B,"HR",C:C,"Present")',
explanation:"Counts records matching multiple conditions.",
example:"Present HR Employees"
},

{
title:"SUMIF",
keywords:[
"sumif","conditional sum","sum condition","sum criteria","sum matching",
"excel sumif","sales total","department total","criteria sum","filtered sum"
],
formula:'=SUMIF(B:B,"HR",D:D)',
explanation:"Adds values that match one condition.",
example:"HR Total Salary"
},

{
title:"SUMIFS",
keywords:[
"sumifs","multiple condition sum","conditional total","criteria sum","sum multiple",
"excel sumifs","region sales","department salary","multi criteria sum","filtered total"
],
formula:'=SUMIFS(D:D,B:B,"HR",C:C,"Male")',
explanation:"Adds values matching multiple conditions.",
example:"Male HR Salary"
},

{
title:"AVERAGEIF",
keywords:[
"averageif","conditional average","average condition","criteria average","filtered average",
"excel averageif","department average","salary average","average formula","average records"
],
formula:'=AVERAGEIF(B:B,"HR",D:D)',
explanation:"Returns the average matching one condition.",
example:"Average HR Salary"
},

{
title:"AVERAGEIFS",
keywords:[
"averageifs","multiple condition average","criteria average","filtered average","average multiple",
"excel averageifs","department salary","multi criteria average","salary analysis","average formula"
],
formula:'=AVERAGEIFS(D:D,B:B,"HR",C:C,"Female")',
explanation:"Returns the average matching multiple conditions.",
example:"Average Female HR Salary"
},

{
title:"MAXIFS",
keywords:[
"maxifs","maximum with criteria","highest value condition","criteria max","conditional maximum",
"excel maxifs","highest salary","department maximum","filtered max","maximum formula"
],
formula:'=MAXIFS(D:D,B:B,"HR")',
explanation:"Returns the maximum value matching criteria.",
example:"Highest HR Salary"
},

{
title:"MINIFS",
keywords:[
"minifs","minimum with criteria","lowest value condition","criteria minimum","conditional minimum",
"excel minifs","lowest salary","department minimum","filtered minimum","minimum formula"
],
formula:'=MINIFS(D:D,B:B,"HR")',
explanation:"Returns the minimum value matching criteria.",
example:"Lowest HR Salary"
},

{
title:"SUBTOTAL",
keywords:[
"subtotal","filtered total","subtotal formula","visible cells","subtotal sum",
"excel subtotal","filter total","subtotal calculation","subtotal values","filtered records"
],
formula:'=SUBTOTAL(9,D2:D100)',
explanation:"Calculates totals while ignoring filtered rows.",
example:"Filtered Salary Total"
},

{
title:"AGGREGATE",
keywords:[
"aggregate","aggregate formula","ignore errors","filtered calculation","advanced subtotal",
"excel aggregate","aggregate function","visible rows","error handling","advanced calculation"
],
formula:'=AGGREGATE(9,5,D2:D100)',
explanation:"Performs advanced calculations while ignoring errors or hidden rows.",
example:"Filtered Total"
},

{
title:"FILTER BY CRITERIA",
keywords:[
"filter criteria","dynamic filter","filter department","filter records","criteria filter",
"excel filter criteria","employee filter","dynamic list","filter data","filtered employees"
],
formula:'=FILTER(A2:D100,B2:B100="HR")',
explanation:"Returns only records matching criteria.",
example:"HR Employee List"
},

{
title:"UNIQUE VALUES",
keywords:[
"unique values","distinct list","remove duplicates","unique departments","distinct records",
"excel unique","unique formula","duplicate removal","dynamic unique","unique employees"
],
formula:'=UNIQUE(B2:B100)',
explanation:"Returns distinct values from a range.",
example:"Department List"
},

{
title:"SORT DATABASE",
keywords:[
"sort database","sort records","database sorting","employee sorting","salary sorting",
"excel sort","dynamic sort","sort formula","ascending data","sorted records"
],
formula:'=SORT(A2:D100,4,-1)',
explanation:"Sorts records by the selected column.",
example:"Highest Salary First"
},

{
title:"ISBLANK",
keywords:[
"isblank","blank","empty cell","cell blank","check blank",
"empty","blank value","is blank","blank formula","missing value"
],
formula:'=ISBLANK(A2)',
explanation:"Checks whether a cell is empty.",
example:"TRUE if A2 is blank"
},

{
title:"ISNUMBER",
keywords:[
"isnumber","number","numeric","check number","number value",
"numeric value","is numeric","contains number","number formula","digits"
],
formula:'=ISNUMBER(A2)',
explanation:"Returns TRUE if the value is numeric.",
example:"100 → TRUE"
},

{
title:"ISTEXT",
keywords:[
"istext","text","string","check text","text value",
"contains text","is string","text formula","letters","character"
],
formula:'=ISTEXT(A2)',
explanation:"Returns TRUE if the value is text.",
example:"Rahul → TRUE"
},

{
title:"ISLOGICAL",
keywords:[
"islogical","logical","true false","boolean","check logical",
"boolean value","logical value","true","false","logical formula"
],
formula:'=ISLOGICAL(A2)',
explanation:"Checks whether the value is TRUE or FALSE.",
example:"TRUE → TRUE"
},

{
title:"ISERROR",
keywords:[
"iserror","error","excel error","check error","formula error",
"value error","cell error","error handling","error formula","is error"
],
formula:'=ISERROR(A2)',
explanation:"Returns TRUE if any error exists.",
example:"#VALUE! → TRUE"
},

{
title:"ISERR",
keywords:[
"iserr","error except na","check error","formula error","excel error",
"is err","error formula","cell error","error check","non na error"
],
formula:'=ISERR(A2)',
explanation:"Returns TRUE for all errors except #N/A.",
example:"#VALUE! → TRUE"
},

{
title:"ISNA",
keywords:[
"isna","na error","not available","lookup error","check na",
"n a","is na","na formula","lookup failed","xlookup error"
],
formula:'=ISNA(A2)',
explanation:"Checks whether the value is #N/A.",
example:"#N/A → TRUE"
},

{
title:"ISNONTEXT",
keywords:[
"isnontext","non text","not text","check non text","non string",
"numeric or blank","non text formula","non characters","value type","not string"
],
formula:'=ISNONTEXT(A2)',
explanation:"Returns TRUE if the value is not text.",
example:"150 → TRUE"
},

{
title:"ISEVEN",
keywords:[
"iseven","even","even number","check even","even formula",
"multiple of 2","even value","number even","is even","even check"
],
formula:'=ISEVEN(A2)',
explanation:"Checks whether a number is even.",
example:"20 → TRUE"
},

{
title:"ISODD",
keywords:[
"isodd","odd","odd number","check odd","odd formula",
"odd value","number odd","odd check","is odd","not even"
],
formula:'=ISODD(A2)',
explanation:"Checks whether a number is odd.",
example:"15 → TRUE"
},

{
title:"ISFORMULA",
keywords:[
"isformula","formula","check formula","cell formula","contains formula",
"formula exists","formula check","excel formula","formula detection","is formula"
],
formula:'=ISFORMULA(A2)',
explanation:"Checks whether a cell contains a formula.",
example:"TRUE"
},

{
title:"ERROR.TYPE",
keywords:[
"error type","errortype","error number","error code","excel error code",
"identify error","error formula","error id","formula error type","error"
],
formula:'=ERROR.TYPE(A2)',
explanation:"Returns the numeric code of an error.",
example:"#DIV/0! → 2"
},

{
title:"TYPE",
keywords:[
"type","data type","cell type","value type","check data type",
"text or number","type formula","excel type","datatype","type check"
],
formula:'=TYPE(A2)',
explanation:"Returns the data type of a value.",
example:"1 = Number"
},

{
title:"CELL",
keywords:[
"cell","cell information","cell address","cell format","cell type",
"cell details","excel cell","cell info","cell formula","address"
],
formula:'=CELL("address",A2)',
explanation:"Returns information about a cell.",
example:"$A$2"
},

{
title:"INFO",
keywords:[
"info","system info","excel info","environment","directory",
"os","excel version","info formula","system","workspace"
],
formula:'=INFO("directory")',
explanation:"Returns information about the current environment.",
example:"Current Folder"
},

{
title:"N",
keywords:[
"n","convert number","number conversion","text to zero","numeric value",
"n function","excel n","convert value","number only","n formula"
],
formula:'=N(A2)',
explanation:"Converts values into numbers where possible.",
example:"TRUE → 1"
},

{
title:"NA",
keywords:[
"na","not available","return na","na formula","missing value",
"lookup missing","excel na","error value","return error","n a function"
],
formula:'=NA()',
explanation:"Returns the #N/A error.",
example:"#N/A"
},

{
title:"SHEET",
keywords:[
"sheet","sheet number","worksheet number","current sheet","sheet id",
"excel sheet","sheet formula","sheet index","worksheet","sheet info"
],
formula:'=SHEET()',
explanation:"Returns the sheet number.",
example:"1"
},

{
title:"SHEETS",
keywords:[
"sheets","total sheets","worksheet count","sheet count","number of sheets",
"excel sheets","count worksheets","sheet total","all sheets","worksheet total"
],
formula:'=SHEETS()',
explanation:"Returns the total number of worksheets.",
example:"5"
},

{
title:"FORMULATEXT",
keywords:[
"formulatext","show formula","display formula","formula text","formula display",
"formula content","excel formula text","formula viewer","show function","formula string"
],
formula:'=FORMULATEXT(A2)',
explanation:"Returns the formula from a referenced cell.",
example:"=SUM(A1:A10)"
},

{
title:"GET CELL ADDRESS",
keywords:[
"cell address","address","current address","reference","cell reference",
"excel address","cell location","address formula","worksheet address","cell name"
],
formula:'=ADDRESS(ROW(A2),COLUMN(A2))',
explanation:"Returns the address of a cell.",
example:"$A$2"
},

{
title:"CURRENT ROW",
keywords:[
"current row","row number","which row","row id","current record",
"excel row","row formula","record number","line number","current position"
],
formula:'=ROW()',
explanation:"Returns the current row number.",
example:"15"
},

{
title:"CURRENT COLUMN",
keywords:[
"current column","column number","which column","column id","column position",
"excel column","column formula","current field","column index","position"
],
formula:'=COLUMN()',
explanation:"Returns the current column number.",
example:"3"
},

{
title:"ROWS",
keywords:[
"rows","count rows","number of rows","row count","total rows",
"excel rows","rows formula","table rows","count records","row total"
],
formula:'=ROWS(A2:A100)',
explanation:"Counts the number of rows.",
example:"99"
},

{
title:"COLUMNS",
keywords:[
"columns","count columns","number of columns","column count","total columns",
"excel columns","columns formula","table columns","field count","column total"
],
formula:'=COLUMNS(A:F)',
explanation:"Counts the number of columns.",
example:"6"
},

{
title:"PMT",
keywords:[
"pmt","loan emi","emi","monthly emi","loan payment",
"emi calculator","home loan","car loan","monthly payment","installment"
],
formula:'=PMT(8%/12,60,-500000)',
explanation:"Calculates the monthly loan payment.",
example:"5-Year Home Loan EMI"
},

{
title:"IPMT",
keywords:[
"ipmt","interest payment","loan interest","emi interest","interest amount",
"loan schedule","interest component","monthly interest","finance","interest formula"
],
formula:'=IPMT(8%/12,1,60,-500000)',
explanation:"Returns the interest portion of a loan payment.",
example:"First Month Interest"
},

{
title:"PPMT",
keywords:[
"ppmt","principal payment","loan principal","emi principal","principal amount",
"loan schedule","principal component","monthly principal","finance","principal formula"
],
formula:'=PPMT(8%/12,1,60,-500000)',
explanation:"Returns the principal portion of a loan payment.",
example:"First Month Principal"
},

{
title:"PV",
keywords:[
"pv","present value","loan value","investment value","current value",
"discounted value","financial value","pv formula","investment","finance"
],
formula:'=PV(8%/12,60,-10000)',
explanation:"Returns the present value of an investment.",
example:"Current Investment Value"
},

{
title:"FV",
keywords:[
"fv","future value","investment growth","future investment","maturity value",
"compound interest","future amount","financial formula","finance","investment"
],
formula:'=FV(8%/12,60,-10000)',
explanation:"Calculates the future value of an investment.",
example:"Investment After 5 Years"
},

{
title:"NPV",
keywords:[
"npv","net present value","investment analysis","project evaluation","cash flow",
"discounted cash flow","finance","npv formula","investment return","valuation"
],
formula:'=NPV(10%,B2:B10)',
explanation:"Calculates the net present value of cash flows.",
example:"Project Profitability"
},

{
title:"IRR",
keywords:[
"irr","internal rate of return","investment return","project return","cash flow return",
"finance","irr formula","rate of return","investment analysis","return percentage"
],
formula:'=IRR(B2:B10)',
explanation:"Returns the internal rate of return.",
example:"Investment IRR"
},

{
title:"XIRR",
keywords:[
"xirr","irregular cash flow","cash flow dates","investment xirr","return with dates",
"finance","annual return","xirr formula","investment analysis","cash flow"
],
formula:'=XIRR(B2:B10,A2:A10)',
explanation:"Returns IRR for irregular cash flow dates.",
example:"Annual Investment Return"
},

{
title:"XNPV",
keywords:[
"xnpv","npv with dates","discount cash flow","investment dates","finance npv",
"xnpv formula","cash flow analysis","valuation","project analysis","investment"
],
formula:'=XNPV(10%,B2:B10,A2:A10)',
explanation:"Calculates NPV using actual dates.",
example:"Project Valuation"
},

{
title:"RATE",
keywords:[
"rate","interest rate","loan rate","investment rate","financial rate",
"calculate rate","loan interest","rate formula","finance","interest"
],
formula:'=RATE(60,-10000,500000)',
explanation:"Calculates the interest rate.",
example:"Loan Interest Rate"
},

{
title:"NPER",
keywords:[
"nper","loan tenure","loan period","number of payments","payment periods",
"finance","loan duration","emi period","nper formula","investment duration"
],
formula:'=NPER(8%/12,-10000,500000)',
explanation:"Returns the number of payment periods.",
example:"Loan Duration"
},

{
title:"SLN",
keywords:[
"sln","straight line depreciation","asset depreciation","depreciation","fixed asset",
"finance","accounting","depreciation formula","asset value","book value"
],
formula:'=SLN(100000,10000,5)',
explanation:"Calculates straight-line depreciation.",
example:"Annual Depreciation"
},

{
title:"DB",
keywords:[
"db","declining balance","asset depreciation","depreciation db","book value",
"finance","fixed asset","declining depreciation","accounting","db formula"
],
formula:'=DB(100000,10000,5,1)',
explanation:"Calculates declining balance depreciation.",
example:"Year 1 Depreciation"
},

{
title:"DDB",
keywords:[
"ddb","double declining balance","accelerated depreciation","asset depreciation","book value",
"finance","ddb formula","accounting","depreciation method","fixed asset"
],
formula:'=DDB(100000,10000,5,1)',
explanation:"Calculates double declining balance depreciation.",
example:"Accelerated Depreciation"
},

{
title:"SYD",
keywords:[
"syd","sum of years digits","asset depreciation","depreciation syd","finance",
"accounting","book value","asset value","depreciation formula","fixed asset"
],
formula:'=SYD(100000,10000,5,1)',
explanation:"Calculates Sum-of-Years' Digits depreciation.",
example:"Year 1 Depreciation"
},

{
title:"CUMIPMT",
keywords:[
"cumipmt","total interest","loan interest total","interest paid","loan schedule",
"finance","cumulative interest","interest calculation","emi","loan"
],
formula:'=CUMIPMT(8%/12,60,500000,1,12,0)',
explanation:"Returns cumulative interest paid.",
example:"First Year Interest"
},

{
title:"CUMPRINC",
keywords:[
"cumprinc","principal paid","loan principal total","principal payment","loan schedule",
"finance","cumulative principal","emi","loan","principal calculation"
],
formula:'=CUMPRINC(8%/12,60,500000,1,12,0)',
explanation:"Returns cumulative principal paid.",
example:"First Year Principal"
},

{
title:"PRICE",
keywords:[
"price","bond price","security price","bond valuation","investment price",
"finance","bond","price formula","financial asset","market value"
],
formula:'=PRICE(DATE(2025,1,1),DATE(2030,1,1),5%,6%,100,2,0)',
explanation:"Returns the price of a bond.",
example:"Bond Market Price"
},

{
title:"YIELD",
keywords:[
"yield","bond yield","investment yield","return on bond","finance",
"bond return","yield formula","financial asset","coupon yield","market yield"
],
formula:'=YIELD(DATE(2025,1,1),DATE(2030,1,1),5%,95,100,2,0)',
explanation:"Returns the bond yield.",
example:"Annual Bond Yield"
},

{
title:"DISC",
keywords:[
"disc","discount security","treasury bill","discount rate","finance",
"discount formula","investment","bond discount","security","disc function"
],
formula:'=DISC(DATE(2025,1,1),DATE(2025,7,1),97,100,0)',
explanation:"Returns the discount rate of a security.",
example:"Treasury Bill Discount"
},

{
title:"DURATION",
keywords:[
"duration","bond duration","investment duration","macaulay duration","finance",
"bond analysis","duration formula","fixed income","portfolio","bond risk"
],
formula:'=DURATION(DATE(2025,1,1),DATE(2030,1,1),5%,6%,2,0)',
explanation:"Returns the annual duration of a bond.",
example:"Bond Duration"
},

{
title:"MDURATION",
keywords:[
"mduration","modified duration","bond risk","interest risk","finance",
"bond duration","modified bond duration","portfolio","investment","fixed income"
],
formula:'=MDURATION(DATE(2025,1,1),DATE(2030,1,1),5%,6%,2,0)',
explanation:"Returns the modified duration of a bond.",
example:"Interest Rate Risk"
},

{
title:"ACCRINT",
keywords:[
"accrint","accrued interest","bond interest","interest earned","coupon interest",
"finance","bond","accrual","investment","fixed income"
],
formula:'=ACCRINT(DATE(2025,1,1),DATE(2025,7,1),DATE(2025,12,31),5%,1000,2)',
explanation:"Calculates accrued interest.",
example:"Bond Interest Earned"
},

{
title:"EFFECT",
keywords:[
"effective interest","effective rate","effective annual rate","ear","finance",
"interest rate","effective formula","loan rate","investment","compound interest"
],
formula:'=EFFECT(8%,12)',
explanation:"Calculates the effective annual interest rate.",
example:"8.30%"
},

{
title:"NOMINAL",
keywords:[
"nominal","nominal rate","annual interest","nominal interest","finance",
"interest formula","loan rate","investment","nominal formula","interest"
],
formula:'=NOMINAL(8.3%,12)',
explanation:"Calculates the nominal annual interest rate.",
example:"8.00%"
},

{
title:"MEDIAN",
keywords:[
"median","middle value","median formula","statistics","middle number",
"center value","median calculation","data analysis","statistical median","middle"
],
formula:'=MEDIAN(A2:A100)',
explanation:"Returns the middle value of a dataset.",
example:"Median Salary"
},

{
title:"MODE.SNGL",
keywords:[
"mode","most frequent","common value","mode single","mode formula",
"statistics","frequent value","most repeated","mode excel","popular value"
],
formula:'=MODE.SNGL(A2:A100)',
explanation:"Returns the most frequently occurring value.",
example:"Most Common Salary"
},

{
title:"MODE.MULT",
keywords:[
"mode multiple","multiple mode","repeated values","common values","mode mult",
"statistics","duplicate values","most frequent","mode formula","multiple results"
],
formula:'=MODE.MULT(A2:A100)',
explanation:"Returns multiple most frequent values.",
example:"Multiple Common Scores"
},

{
title:"STDEV.S",
keywords:[
"stdev","standard deviation","sample deviation","stdev sample","variation",
"statistics","spread","salary deviation","data spread","stdev.s"
],
formula:'=STDEV.S(A2:A100)',
explanation:"Calculates sample standard deviation.",
example:"Salary Variation"
},

{
title:"STDEV.P",
keywords:[
"stdevp","population deviation","standard deviation population","stdev population",
"statistics","variation","spread","data analysis","stdev.p","population"
],
formula:'=STDEV.P(A2:A100)',
explanation:"Calculates population standard deviation.",
example:"Company Salary Deviation"
},

{
title:"VAR.S",
keywords:[
"variance","sample variance","var sample","data variance","variation",
"statistics","spread","var.s","sample statistics","variance formula"
],
formula:'=VAR.S(A2:A100)',
explanation:"Returns sample variance.",
example:"Salary Variance"
},

{
title:"VAR.P",
keywords:[
"population variance","var population","variance population","var.p","statistics",
"variation","spread","population statistics","variance formula","data variance"
],
formula:'=VAR.P(A2:A100)',
explanation:"Returns population variance.",
example:"Company Variance"
},

{
title:"PERCENTILE.INC",
keywords:[
"percentile","90 percentile","percentile inc","ranking","statistics",
"top percentile","data analysis","score percentile","percentile formula","distribution"
],
formula:'=PERCENTILE.INC(A2:A100,0.9)',
explanation:"Returns the specified percentile.",
example:"90th Percentile Salary"
},

{
title:"QUARTILE.INC",
keywords:[
"quartile","quartiles","first quartile","third quartile","quartile inc",
"statistics","q1","q3","quartile formula","distribution"
],
formula:'=QUARTILE.INC(A2:A100,1)',
explanation:"Returns the specified quartile.",
example:"First Quartile"
},

{
title:"PERCENTRANK.INC",
keywords:[
"percent rank","ranking percentage","percentile rank","rank percentage",
"statistics","relative rank","score rank","percentrank","ranking","analysis"
],
formula:'=PERCENTRANK.INC(A2:A100,A2)',
explanation:"Returns the percentage rank of a value.",
example:"Employee Performance Rank"
},

{
title:"CORREL",
keywords:[
"correlation","correl","relationship","compare datasets","correlation coefficient",
"statistics","data relationship","analysis","trend","correlation analysis"
],
formula:'=CORREL(A2:A100,B2:B100)',
explanation:"Returns the correlation coefficient.",
example:"Sales vs Profit"
},

{
title:"COVARIANCE.S",
keywords:[
"covariance","sample covariance","covariance sample","relationship",
"statistics","covariance.s","data analysis","covariance formula","sample","analysis"
],
formula:'=COVARIANCE.S(A2:A100,B2:B100)',
explanation:"Returns sample covariance.",
example:"Sales & Profit"
},

{
title:"COVARIANCE.P",
keywords:[
"population covariance","covariance population","covariance.p","statistics",
"relationship","analysis","covariance formula","population","data analysis","covariance"
],
formula:'=COVARIANCE.P(A2:A100,B2:B100)',
explanation:"Returns population covariance.",
example:"Population Covariance"
},

{
title:"FORECAST.LINEAR",
keywords:[
"forecast","forecast linear","prediction","predict value","future value",
"linear forecast","excel forecast","forecast formula","trend prediction","projection"
],
formula:'=FORECAST.LINEAR(2026,B2:B10,A2:A10)',
explanation:"Predicts a future value using linear regression.",
example:"Next Year Sales"
},

{
title:"TREND",
keywords:[
"trend","trend analysis","future trend","sales trend","forecast trend",
"excel trend","trend formula","growth trend","prediction","trendline"
],
formula:'=TREND(B2:B10,A2:A10,A11)',
explanation:"Returns values along a linear trend.",
example:"Predicted Sales"
},

{
title:"LINEST",
keywords:[
"linest","linear regression","regression","trendline","line equation",
"statistics","forecast","data analysis","regression analysis","line fit"
],
formula:'=LINEST(B2:B10,A2:A10)',
explanation:"Calculates linear regression statistics.",
example:"Regression Analysis"
},

{
title:"SLOPE",
keywords:[
"slope","trend slope","regression slope","line slope","growth rate",
"statistics","trend","analysis","slope formula","linear regression"
],
formula:'=SLOPE(B2:B10,A2:A10)',
explanation:"Returns the slope of a regression line.",
example:"Sales Growth Rate"
},

{
title:"INTERCEPT",
keywords:[
"intercept","y intercept","regression intercept","line intercept","statistics",
"trend","analysis","intercept formula","linear regression","forecast"
],
formula:'=INTERCEPT(B2:B10,A2:A10)',
explanation:"Returns the intercept of a regression line.",
example:"Regression Intercept"
},

{
title:"RSQ",
keywords:[
"rsq","r square","coefficient determination","correlation square","statistics",
"regression","analysis","goodness fit","r2","trend analysis"
],
formula:'=RSQ(B2:B10,A2:A10)',
explanation:"Returns the R-squared value.",
example:"Regression Accuracy"
},

{
title:"STEYX",
keywords:[
"steyx","standard error","regression error","prediction error","statistics",
"forecast","analysis","standard error estimate","regression","error analysis"
],
formula:'=STEYX(B2:B10,A2:A10)',
explanation:"Returns the standard error of prediction.",
example:"Forecast Accuracy"
},

{
title:"GEOMEAN",
keywords:[
"geometric mean","geomean","average growth","compound average","statistics",
"growth rate","investment return","mean","geometric average","analysis"
],
formula:'=GEOMEAN(A2:A100)',
explanation:"Returns the geometric mean.",
example:"Average Investment Growth"
},

{
title:"HARMEAN",
keywords:[
"harmonic mean","harmean","harmonic average","statistics","mean",
"average rate","speed average","harmonic","analysis","formula"
],
formula:'=HARMEAN(A2:A100)',
explanation:"Returns the harmonic mean.",
example:"Average Speed"
},

{
title:"SKEW",
keywords:[
"skew","data skewness","distribution","statistics","asymmetry",
"skewness","data analysis","distribution shape","skew formula","analysis"
],
formula:'=SKEW(A2:A100)',
explanation:"Measures the asymmetry of a dataset.",
example:"Salary Distribution"
},

{
title:"KURT",
keywords:[
"kurt","kurtosis","distribution peak","statistics","distribution shape",
"kurtosis formula","analysis","peakedness","data analysis","kurt"
],
formula:'=KURT(A2:A100)',
explanation:"Returns the kurtosis of a dataset.",
example:"Distribution Shape"
},

{
title:"FREQUENCY",
keywords:[
"frequency","frequency distribution","histogram","count ranges","bins",
"statistics","frequency formula","distribution","data grouping","analysis"
],
formula:'=FREQUENCY(A2:A100,D2:D10)',
explanation:"Calculates how often values occur within ranges.",
example:"Marks Distribution"
},

{
title:"CONVERT",
keywords:[
"convert","unit conversion","meters to feet","kg to pound","temperature conversion",
"excel convert","measurement","convert units","length conversion","weight conversion"
],
formula:'=CONVERT(10,"m","ft")',
explanation:"Converts a number from one measurement unit to another.",
example:"10 Meters → 32.81 Feet"
},

{
title:"DELTA",
keywords:[
"delta","compare numbers","equal numbers","check equality","engineering delta",
"excel delta","compare values","same number","delta function","equal check"
],
formula:'=DELTA(A2,B2)',
explanation:"Returns 1 if two numbers are equal; otherwise returns 0.",
example:"100 = 100 → 1"
},

{
title:"GESTEP",
keywords:[
"gestep","greater than","step function","compare threshold","engineering",
"excel gestep","number comparison","greater check","step formula","threshold"
],
formula:'=GESTEP(A2,100)',
explanation:"Returns 1 if a number is greater than or equal to the step value.",
example:"120 ≥ 100 → 1"
},

{
title:"BIN2DEC",
keywords:[
"binary to decimal","bin2dec","binary conversion","binary number","convert binary",
"excel binary","decimal conversion","binary decimal","engineering","number system"
],
formula:'=BIN2DEC("1010")',
explanation:"Converts a binary number to decimal.",
example:"1010 → 10"
},

{
title:"DEC2BIN",
keywords:[
"decimal to binary","dec2bin","binary converter","decimal conversion","convert decimal",
"excel binary","binary format","number conversion","engineering","binary"
],
formula:'=DEC2BIN(10)',
explanation:"Converts a decimal number to binary.",
example:"10 → 1010"
},

{
title:"BIN2HEX",
keywords:[
"binary to hexadecimal","bin2hex","binary hex","convert binary","hex conversion",
"engineering","excel binary","hexadecimal","binary format","hex"
],
formula:'=BIN2HEX("1111")',
explanation:"Converts binary to hexadecimal.",
example:"1111 → F"
},

{
title:"HEX2BIN",
keywords:[
"hex to binary","hex2bin","hexadecimal binary","convert hex","engineering",
"excel hex","binary conversion","hex number","number system","binary"
],
formula:'=HEX2BIN("F")',
explanation:"Converts hexadecimal to binary.",
example:"F → 1111"
},

{
title:"DEC2HEX",
keywords:[
"decimal to hex","dec2hex","hexadecimal","convert decimal","engineering",
"excel hex","hex conversion","number system","hex format","decimal conversion"
],
formula:'=DEC2HEX(255)',
explanation:"Converts decimal to hexadecimal.",
example:"255 → FF"
},

{
title:"HEX2DEC",
keywords:[
"hex to decimal","hex2dec","convert hex","hexadecimal conversion","engineering",
"excel hex","decimal format","number system","hex decimal","conversion"
],
formula:'=HEX2DEC("FF")',
explanation:"Converts hexadecimal to decimal.",
example:"FF → 255"
},

{
title:"BIN2OCT",
keywords:[
"binary to octal","bin2oct","binary octal","convert binary","engineering",
"octal conversion","excel engineering","number systems","binary","octal"
],
formula:'=BIN2OCT("11111111")',
explanation:"Converts binary to octal.",
example:"11111111 → 377"
},

{
title:"OCT2BIN",
keywords:[
"octal to binary","oct2bin","convert octal","binary conversion","engineering",
"excel octal","number system","octal binary","binary","octal"
],
formula:'=OCT2BIN("377")',
explanation:"Converts octal to binary.",
example:"377 → 11111111"
},

{
title:"DEC2OCT",
keywords:[
"decimal to octal","dec2oct","convert decimal","octal conversion","engineering",
"excel octal","number system","decimal octal","conversion","octal"
],
formula:'=DEC2OCT(255)',
explanation:"Converts decimal to octal.",
example:"255 → 377"
},

{
title:"OCT2DEC",
keywords:[
"octal to decimal","oct2dec","convert octal","engineering","octal decimal",
"excel octal","number conversion","octal","decimal","number system"
],
formula:'=OCT2DEC("377")',
explanation:"Converts octal to decimal.",
example:"377 → 255"
},

{
title:"HEX2OCT",
keywords:[
"hex to octal","hex2oct","hexadecimal octal","engineering","convert hex",
"octal conversion","excel engineering","hex","octal","number system"
],
formula:'=HEX2OCT("FF")',
explanation:"Converts hexadecimal to octal.",
example:"FF → 377"
},

{
title:"OCT2HEX",
keywords:[
"octal to hexadecimal","oct2hex","octal hex","engineering","convert octal",
"hex conversion","excel engineering","hexadecimal","octal","number system"
],
formula:'=OCT2HEX("377")',
explanation:"Converts octal to hexadecimal.",
example:"377 → FF"
},

{
title:"COMPLEX",
keywords:[
"complex","complex number","imaginary number","engineering math","complex formula",
"excel complex","real imaginary","engineering","math","complex values"
],
formula:'=COMPLEX(3,4)',
explanation:"Creates a complex number.",
example:"3+4i"
},

{
title:"IMABS",
keywords:[
"complex absolute","imabs","absolute complex","engineering","complex magnitude",
"excel complex","imaginary","math","engineering function","absolute"
],
formula:'=IMABS("3+4i")',
explanation:"Returns the absolute value of a complex number.",
example:"5"
},

{
title:"IMAGINARY",
keywords:[
"imaginary part","imaginary","complex imaginary","engineering","excel complex",
"imaginary value","complex number","math","engineering formula","imaginary component"
],
formula:'=IMAGINARY("3+4i")',
explanation:"Returns the imaginary coefficient.",
example:"4"
},

{
title:"IMREAL",
keywords:[
"real part","imreal","complex real","engineering","real value",
"excel complex","complex number","math","engineering","real component"
],
formula:'=IMREAL("3+4i")',
explanation:"Returns the real coefficient.",
example:"3"
},

{
title:"IMSUM",
keywords:[
"complex sum","imsum","add complex","engineering","complex addition",
"excel complex","imaginary","math","complex formula","sum"
],
formula:'=IMSUM("2+3i","4+5i")',
explanation:"Adds complex numbers.",
example:"6+8i"
},

{
title:"IMSUB",
keywords:[
"complex subtraction","imsub","subtract complex","engineering","complex minus",
"excel complex","imaginary","math","complex calculation","subtract"
],
formula:'=IMSUB("6+5i","2+1i")',
explanation:"Subtracts complex numbers.",
example:"4+4i"
},

{
title:"IMPRODUCT",
keywords:[
"complex multiplication","improduct","multiply complex","engineering","complex product",
"excel complex","imaginary","math","multiply","complex math"
],
formula:'=IMPRODUCT("2+3i","4+5i")',
explanation:"Multiplies complex numbers.",
example:"-7+22i"
},

{
title:"IMDIV",
keywords:[
"complex division","imdiv","divide complex","engineering","complex quotient",
"excel complex","imaginary","math","division","complex calculation"
],
formula:'=IMDIV("6+8i","2+2i")',
explanation:"Divides one complex number by another.",
example:"3.5+0.5i"
},

{
title:"IMSQRT",
keywords:[
"complex square root","imsqrt","sqrt complex","engineering","complex root",
"excel complex","imaginary","math","square root","complex"
],
formula:'=IMSQRT("16")',
explanation:"Returns the square root of a complex number.",
example:"4"
},

{
title:"IMPOWER",
keywords:[
"complex power","impower","power complex","engineering","complex exponent",
"excel complex","imaginary","math","power","complex formula"
],
formula:'=IMPOWER("2+3i",2)',
explanation:"Raises a complex number to a power.",
example:"(-5+12i)"
},

{
title:"PI",
keywords:[
"pi","pi value","3.14159","circle pi","mathematical constant",
"excel pi","π","geometry","math constant","pi function"
],
formula:'=PI()',
explanation:"Returns the value of Pi.",
example:"3.141593"
},

{
title:"RADIANS",
keywords:[
"radians","degree to radian","convert degree","angle conversion","math",
"excel radians","angle","geometry","radian formula","degree conversion"
],
formula:'=RADIANS(180)',
explanation:"Converts degrees into radians.",
example:"180° → 3.14159"
},

{
title:"DEGREES",
keywords:[
"degrees","radian to degree","convert radians","angle conversion","math",
"excel degrees","geometry","degree formula","radian conversion","angle"
],
formula:'=DEGREES(PI())',
explanation:"Converts radians into degrees.",
example:"π → 180°"
},

{
title:"SIN",
keywords:[
"sin","sine","trigonometry","sin angle","math",
"excel sin","triangle","wave","engineering","sine formula"
],
formula:'=SIN(RADIANS(30))',
explanation:"Returns the sine of an angle.",
example:"30° → 0.5"
},

{
title:"COS",
keywords:[
"cos","cosine","trigonometry","cos angle","math",
"excel cos","triangle","wave","engineering","cosine formula"
],
formula:'=COS(RADIANS(60))',
explanation:"Returns the cosine of an angle.",
example:"60° → 0.5"
},

{
title:"TAN",
keywords:[
"tan","tangent","trigonometry","tan angle","math",
"excel tan","triangle","engineering","tangent formula","angle"
],
formula:'=TAN(RADIANS(45))',
explanation:"Returns the tangent of an angle.",
example:"45° → 1"
},

{
title:"ASIN",
keywords:[
"asin","inverse sine","arc sine","trigonometry","math",
"excel asin","inverse trig","angle","engineering","arcsin"
],
formula:'=DEGREES(ASIN(0.5))',
explanation:"Returns the inverse sine.",
example:"0.5 → 30°"
},

{
title:"ACOS",
keywords:[
"acos","inverse cosine","arc cosine","trigonometry","math",
"excel acos","inverse trig","engineering","arccos","angle"
],
formula:'=DEGREES(ACOS(0.5))',
explanation:"Returns the inverse cosine.",
example:"0.5 → 60°"
},

{
title:"ATAN",
keywords:[
"atan","inverse tangent","arc tangent","trigonometry","math",
"excel atan","inverse trig","engineering","arctan","angle"
],
formula:'=DEGREES(ATAN(1))',
explanation:"Returns the inverse tangent.",
example:"1 → 45°"
},

{
title:"ATAN2",
keywords:[
"atan2","two coordinate angle","x y angle","coordinate angle","engineering",
"excel atan2","cartesian","math","angle calculation","trigonometry"
],
formula:'=DEGREES(ATAN2(10,10))',
explanation:"Returns the angle from X and Y coordinates.",
example:"45°"
},

{
title:"SINH",
keywords:[
"sinh","hyperbolic sine","engineering","math","hyperbolic",
"excel sinh","sine hyperbolic","advanced math","science","formula"
],
formula:'=SINH(1)',
explanation:"Returns the hyperbolic sine.",
example:"1.1752"
},

{
title:"COSH",
keywords:[
"cosh","hyperbolic cosine","engineering","math","hyperbolic",
"excel cosh","advanced math","science","formula","cosine"
],
formula:'=COSH(1)',
explanation:"Returns the hyperbolic cosine.",
example:"1.5431"
},

{
title:"TANH",
keywords:[
"tanh","hyperbolic tangent","engineering","math","hyperbolic",
"excel tanh","advanced math","science","formula","tangent"
],
formula:'=TANH(1)',
explanation:"Returns the hyperbolic tangent.",
example:"0.7616"
},

{
title:"ASINH",
keywords:[
"asinh","inverse hyperbolic sine","engineering","math","hyperbolic inverse",
"excel asinh","science","formula","inverse sinh","advanced math"
],
formula:'=ASINH(2)',
explanation:"Returns the inverse hyperbolic sine.",
example:"1.4436"
},

{
title:"ACOSH",
keywords:[
"acosh","inverse hyperbolic cosine","engineering","math","hyperbolic inverse",
"excel acosh","science","formula","inverse cosh","advanced math"
],
formula:'=ACOSH(5)',
explanation:"Returns the inverse hyperbolic cosine.",
example:"2.2924"
},

{
title:"ATANH",
keywords:[
"atanh","inverse hyperbolic tangent","engineering","math","hyperbolic inverse",
"excel atanh","science","formula","inverse tanh","advanced math"
],
formula:'=ATANH(0.5)',
explanation:"Returns the inverse hyperbolic tangent.",
example:"0.5493"
},

{
title:"COT",
keywords:[
"cot","cotangent","trigonometry","cot angle","math",
"excel cot","engineering","triangle","cotangent formula","angle"
],
formula:'=COT(RADIANS(45))',
explanation:"Returns the cotangent of an angle.",
example:"45° → 1"
},

{
title:"CSC",
keywords:[
"csc","cosecant","trigonometry","math","engineering",
"excel csc","triangle","angle","cosecant formula","advanced trig"
],
formula:'=CSC(RADIANS(30))',
explanation:"Returns the cosecant of an angle.",
example:"30° → 2"
},

{
title:"SEC",
keywords:[
"sec","secant","trigonometry","math","engineering",
"excel sec","triangle","angle","secant formula","advanced trig"
],
formula:'=SEC(RADIANS(60))',
explanation:"Returns the secant of an angle.",
example:"60° → 2"
},

{
title:"ACOT",
keywords:[
"acot","inverse cotangent","arc cotangent","math","engineering",
"excel acot","inverse trig","advanced trig","cotangent","angle"
],
formula:'=DEGREES(ACOT(1))',
explanation:"Returns the inverse cotangent.",
example:"1 → 45°"
},

{
title:"COTH",
keywords:[
"coth","hyperbolic cotangent","engineering","math","advanced trig",
"excel coth","science","formula","hyperbolic","cotangent"
],
formula:'=COTH(2)',
explanation:"Returns the hyperbolic cotangent.",
example:"1.0373"
},

{
title:"CSCH",
keywords:[
"csch","hyperbolic cosecant","engineering","math","advanced trig",
"excel csch","science","formula","hyperbolic","cosecant"
],
formula:'=CSCH(2)',
explanation:"Returns the hyperbolic cosecant.",
example:"0.2757"
},

{
title:"SECH",
keywords:[
"sech","hyperbolic secant","engineering","math","advanced trig",
"excel sech","science","formula","hyperbolic","secant"
],
formula:'=SECH(2)',
explanation:"Returns the hyperbolic secant.",
example:"0.2658"
},

{
title:"MUNIT",
keywords:[
"munit","identity matrix","matrix","engineering","matrix identity",
"excel munit","math","array","matrix function","unit matrix"
],
formula:'=MUNIT(3)',
explanation:"Returns a 3×3 identity matrix.",
example:"Identity Matrix"
},

{
title:"MMULT",
keywords:[
"mmult","matrix multiplication","multiply matrix","matrix product","engineering",
"excel mmult","linear algebra","array math","matrix formula","matrix calculation"
],
formula:'=MMULT(A1:C3,E1:G3)',
explanation:"Returns the matrix product of two arrays.",
example:"Matrix Multiplication"
},

{
title:"GCD",
keywords:[
"gcd","greatest common divisor","highest common factor","hcf","common divisor",
"math","engineering","gcd formula","number theory","greatest factor"
],
formula:'=GCD(A2,B2)',
explanation:"Returns the greatest common divisor of two or more integers.",
example:"24 & 36 → 12"
},

{
title:"LCM",
keywords:[
"lcm","least common multiple","lowest common multiple","common multiple","math",
"engineering","lcm formula","number theory","multiple","smallest multiple"
],
formula:'=LCM(A2,B2)',
explanation:"Returns the least common multiple.",
example:"12 & 18 → 36"
},

{
title:"FACT",
keywords:[
"factorial","fact","n factorial","math","combinatorics",
"engineering","factorial formula","number","statistics","fact function"
],
formula:'=FACT(5)',
explanation:"Returns the factorial of a number.",
example:"5! → 120"
},

{
title:"FACTDOUBLE",
keywords:[
"double factorial","factdouble","math","engineering","double fact",
"number theory","factorial","advanced math","statistics","formula"
],
formula:'=FACTDOUBLE(8)',
explanation:"Returns the double factorial of a number.",
example:"8!! → 384"
},

{
title:"COMBIN",
keywords:[
"combinations","combin","choose","ncr","combination formula",
"statistics","probability","math","selection","engineering"
],
formula:'=COMBIN(10,3)',
explanation:"Returns the number of combinations.",
example:"10C3 → 120"
},

{
title:"COMBINA",
keywords:[
"combina","combinations with repetition","repeated combinations","math","statistics",
"probability","engineering","selection","formula","combination"
],
formula:'=COMBINA(10,3)',
explanation:"Returns combinations with repetition.",
example:"220"
},

{
title:"PERMUT",
keywords:[
"permutation","permut","npr","arrangement","math",
"statistics","engineering","permutation formula","order","selection"
],
formula:'=PERMUT(10,3)',
explanation:"Returns the number of permutations.",
example:"10P3 → 720"
},

{
title:"PERMUTATIONA",
keywords:[
"permutationa","permutation with repetition","arrangement repetition","math","statistics",
"engineering","formula","permutation","selection","advanced math"
],
formula:'=PERMUTATIONA(10,3)',
explanation:"Returns permutations with repetition.",
example:"1000"
},

{
title:"MULTINOMIAL",
keywords:[
"multinomial","multinomial coefficient","statistics","math","probability",
"engineering","formula","advanced math","factorial","combination"
],
formula:'=MULTINOMIAL(2,3,4)',
explanation:"Returns the multinomial coefficient.",
example:"1260"
},

{
title:"MROUND",
keywords:[
"mround","round multiple","nearest multiple","round nearest","math",
"excel rounding","multiple","engineering","round formula","nearest value"
],
formula:'=MROUND(27,5)',
explanation:"Rounds a number to the nearest specified multiple.",
example:"27 → 25"
},

{
title:"EVEN",
keywords:[
"even","round even","nearest even","even number","math",
"engineering","excel even","rounding","formula","number"
],
formula:'=EVEN(7.2)',
explanation:"Rounds a number away from zero to the nearest even integer.",
example:"7.2 → 8"
},

{
title:"ODD",
keywords:[
"odd","round odd","nearest odd","odd number","math",
"engineering","excel odd","rounding","formula","number"
],
formula:'=ODD(6.1)',
explanation:"Rounds a number away from zero to the nearest odd integer.",
example:"6.1 → 7"
},

{
title:"TRUNC",
keywords:[
"trunc","truncate","remove decimals","integer part","math",
"engineering","truncate formula","decimal removal","number","excel trunc"
],
formula:'=TRUNC(123.987,2)',
explanation:"Removes the decimal portion without rounding.",
example:"123.987 → 123.98"
},

{
title:"SIGN",
keywords:[
"sign","positive negative","number sign","math","engineering",
"excel sign","check sign","positive","negative","zero"
],
formula:'=SIGN(A2)',
explanation:"Returns the sign of a number.",
example:"-50 → -1"
},

{
title:"ROMAN",
keywords:[
"roman","roman numeral","convert roman","number to roman","roman format",
"history","math","excel roman","roman conversion","numeral"
],
formula:'=ROMAN(2025)',
explanation:"Converts an Arabic number to Roman numerals.",
example:"2025 → MMXXV"
},

{
title:"ARABIC",
keywords:[
"arabic","roman to number","convert roman","roman conversion","arabic number",
"excel arabic","numeral","history","math","roman numeral"
],
formula:'=ARABIC("MMXXV")',
explanation:"Converts Roman numerals to Arabic numbers.",
example:"MMXXV → 2025"
},

{
title:"BASE",
keywords:[
"base","convert base","binary octal hex","number base","engineering",
"excel base","number conversion","radix","math","base function"
],
formula:'=BASE(255,16)',
explanation:"Converts a decimal number into another base.",
example:"255 → FF"
},

{
title:"DECIMAL",
keywords:[
"decimal","convert decimal","base to decimal","hex decimal","binary decimal",
"engineering","number conversion","excel decimal","radix","math"
],
formula:'=DECIMAL("FF",16)',
explanation:"Converts a number in another base into decimal.",
example:"FF → 255"
},

{
title:"SUMPRODUCT",
keywords:[
"sumproduct","weighted sum","multiply sum","array multiplication","excel analysis",
"financial analysis","weighted average","math","sum product","calculation"
],
formula:'=SUMPRODUCT(A2:A10,B2:B10)',
explanation:"Multiplies corresponding arrays and returns their sum.",
example:"Total Sales Value"
},

{
title:"SUMSQ",
keywords:[
"sumsq","sum squares","square sum","math","statistics",
"engineering","sum of squares","formula","analysis","excel sumsq"
],
formula:'=SUMSQ(A2:A10)',
explanation:"Returns the sum of squares of values.",
example:"5²+6²+7²"
},

{
title:"SUMX2MY2",
keywords:[
"sumx2my2","difference squares","array math","engineering","statistics",
"advanced math","sum difference","formula","excel","matrix"
],
formula:'=SUMX2MY2(A2:A10,B2:B10)',
explanation:"Returns the sum of the difference of squares.",
example:"Array Difference"
},

{
title:"SUMX2PY2",
keywords:[
"sumx2py2","sum squares arrays","array addition","engineering","statistics",
"advanced math","formula","excel","matrix","sum arrays"
],
formula:'=SUMX2PY2(A2:A10,B2:B10)',
explanation:"Returns the sum of the sum of squares.",
example:"Array Square Sum"
},

{
title:"SUMXMY2",
keywords:[
"sumxmy2","square differences","array difference","engineering","statistics",
"advanced math","formula","excel","matrix","difference"
],
formula:'=SUMXMY2(A2:A10,B2:B10)',
explanation:"Returns the sum of squared differences.",
example:"Variance Calculation"
},

{
title:"MDETERM",
keywords:[
"mdeterm","matrix determinant","determinant","linear algebra","matrix",
"engineering","excel matrix","advanced math","determinant formula","matrix calculation"
],
formula:'=MDETERM(A1:C3)',
explanation:"Returns the determinant of a matrix.",
example:"Determinant Value"
},

{
title:"MINVERSE",
keywords:[
"minverse","inverse matrix","matrix inverse","linear algebra","matrix",
"engineering","excel matrix","advanced math","inverse calculation","matrix formula"
],
formula:'=MINVERSE(A1:C3)',
explanation:"Returns the inverse of a matrix.",
example:"Inverse Matrix"
},

{
title:"TAKE",
keywords:[
"take","take rows","take columns","first rows","first columns",
"extract top rows","dynamic array","excel take","take function","top records"
],
formula:'=TAKE(A2:D100,5)',
explanation:"Returns the first rows or columns from an array.",
example:"First 5 Employees"
},

{
title:"DROP",
keywords:[
"drop","remove first rows","skip rows","drop records","remove top rows",
"dynamic array","excel drop","drop function","ignore first rows","exclude rows"
],
formula:'=DROP(A2:D100,5)',
explanation:"Removes specified rows or columns from an array.",
example:"Skip First 5 Records"
},

{
title:"CHOOSEROWS",
keywords:[
"chooserows","select rows","pick rows","specific rows","dynamic array",
"excel chooserows","row selection","extract rows","chosen rows","array rows"
],
formula:'=CHOOSEROWS(A2:D20,1,5,10)',
explanation:"Returns specified rows from an array.",
example:"Rows 1,5 & 10"
},

{
title:"CHOOSECOLS",
keywords:[
"choosecols","select columns","pick columns","specific columns","dynamic array",
"excel choosecols","column selection","extract columns","chosen columns","array columns"
],
formula:'=CHOOSECOLS(A2:F100,1,3,5)',
explanation:"Returns specified columns from an array.",
example:"Employee ID, Department & Salary"
},

{
title:"VSTACK",
keywords:[
"vstack","append tables","combine tables","merge vertically","stack rows",
"dynamic array","excel vstack","combine datasets","append data","merge rows"
],
formula:'=VSTACK(A2:D20,F2:I20)',
explanation:"Stacks multiple arrays vertically.",
example:"Combine Two Employee Lists"
},

{
title:"HSTACK",
keywords:[
"hstack","combine columns","merge horizontally","stack columns","join arrays",
"dynamic array","excel hstack","combine tables","merge columns","append columns"
],
formula:'=HSTACK(A2:B20,D2:E20)',
explanation:"Stacks multiple arrays horizontally.",
example:"Join Employee Details"
},

{
title:"TOCOL",
keywords:[
"tocol","convert column","flatten column","single column","array column",
"dynamic array","excel tocol","flatten array","convert vertical","column array"
],
formula:'=TOCOL(A2:D10)',
explanation:"Converts an array into a single column.",
example:"One Long Column"
},

{
title:"TOROW",
keywords:[
"torow","convert row","flatten row","single row","array row",
"dynamic array","excel torow","flatten array","convert horizontal","row array"
],
formula:'=TOROW(A2:D10)',
explanation:"Converts an array into a single row.",
example:"One Long Row"
},

{
title:"WRAPROWS",
keywords:[
"wraprows","wrap into rows","reshape rows","array reshape","dynamic array",
"excel wraprows","split rows","reshape array","row grouping","wrap function"
],
formula:'=WRAPROWS(A2:A12,3)',
explanation:"Wraps a column into multiple rows.",
example:"3 Values Per Row"
},

{
title:"WRAPCOLS",
keywords:[
"wrapcols","wrap into columns","reshape columns","array reshape","dynamic array",
"excel wrapcols","split columns","reshape array","column grouping","wrap function"
],
formula:'=WRAPCOLS(A2:A12,3)',
explanation:"Wraps a column into multiple columns.",
example:"3 Values Per Column"
},

{
title:"EXPAND",
keywords:[
"expand","expand array","resize array","increase array","dynamic array",
"excel expand","array size","pad array","expand function","fill blanks"
],
formula:'=EXPAND(A2:C5,10,5)',
explanation:"Expands an array to specified dimensions.",
example:"Expand to 10×5"
},

{
title:"LET",
keywords:[
"let","variable","named variable","simplify formula","formula variable",
"excel let","reuse calculation","optimize formula","advanced excel","formula optimization"
],
formula:'=LET(x,A2*10,x+100)',
explanation:"Assigns names to calculations inside a formula.",
example:"Improves Formula Readability"
},

{
title:"LAMBDA",
keywords:[
"lambda","custom function","user defined formula","create function","advanced excel",
"excel lambda","reusable formula","lambda function","custom calculation","formula automation"
],
formula:'=LAMBDA(x,x*10)(5)',
explanation:"Creates reusable custom Excel functions.",
example:"5 → 50"
},

{
title:"MAP",
keywords:[
"map","apply formula","array calculation","loop array","dynamic array",
"excel map","calculate each row","advanced excel","map function","iterate array"
],
formula:'=MAP(A2:A10,LAMBDA(x,x*2))',
explanation:"Applies a calculation to each value in an array.",
example:"Double Every Value"
},

{
title:"REDUCE",
keywords:[
"reduce","accumulate","running calculation","sum array","advanced array",
"excel reduce","aggregate array","lambda reduce","dynamic array","reduce function"
],
formula:'=REDUCE(0,A2:A10,LAMBDA(a,b,a+b))',
explanation:"Reduces an array into a single value.",
example:"Running Total"
},

{
title:"SCAN",
keywords:[
"scan","running total","cumulative total","accumulate values","dynamic array",
"excel scan","running balance","lambda scan","progressive total","array calculation"
],
formula:'=SCAN(0,A2:A10,LAMBDA(a,b,a+b))',
explanation:"Returns cumulative calculations for each element.",
example:"Running Balance"
},

{
title:"BYROW",
keywords:[
"byrow","calculate by row","row calculation","dynamic array","lambda row",
"excel byrow","row wise","apply row","advanced excel","row function"
],
formula:'=BYROW(A2:D10,LAMBDA(r,SUM(r)))',
explanation:"Calculates each row individually.",
example:"Row Totals"
},

{
title:"BYCOL",
keywords:[
"bycol","calculate by column","column calculation","dynamic array","lambda column",
"excel bycol","column wise","apply column","advanced excel","column function"
],
formula:'=BYCOL(A2:D10,LAMBDA(c,SUM(c)))',
explanation:"Calculates each column individually.",
example:"Column Totals"
},

{
title:"MAKEARRAY",
keywords:[
"makearray","generate array","create array","dynamic array","array generator",
"excel makearray","lambda array","advanced excel","array function","matrix"
],
formula:'=MAKEARRAY(5,5,LAMBDA(r,c,r*c))',
explanation:"Creates a custom array using a LAMBDA formula.",
example:"5×5 Multiplication Table"
},

{
title:"ISOMITTED",
keywords:[
"isomitted","optional parameter","lambda optional","advanced excel","lambda parameter",
"excel isomitted","check omitted","custom function","optional argument","lambda"
],
formula:'=ISOMITTED(A2)',
explanation:"Checks whether an optional LAMBDA argument was omitted.",
example:"TRUE/FALSE"
},

{
title:"GROUPBY",
keywords:[
"groupby","group data","summarize data","group records","aggregate data",
"excel groupby","pivot alternative","group function","dynamic summary","data grouping"
],
formula:'=GROUPBY(A2:A100,B2:B100,SUM)',
explanation:"Groups values and returns aggregated results (Microsoft 365).",
example:"Sales by Department"
},

{
title:"PIVOTBY",
keywords:[
"pivotby","pivot table formula","dynamic pivot","pivot function","group summary",
"excel pivotby","pivot analysis","summarize table","pivot formula","data analysis"
],
formula:'=PIVOTBY(A2:A100,B2:B100,C2:C100,SUM)',
explanation:"Creates PivotTable-like summaries using formulas (Microsoft 365).",
example:"Department-wise Sales"
},

{
title:"IMAGE",
keywords:[
"image","display image","picture from url","excel image","online image",
"show image","image function","photo url","dynamic image","product image"
],
formula:'=IMAGE("https://example.com/photo.jpg")',
explanation:"Displays an image directly from a URL.",
example:"Product Photo"
},

{
title:"TEXTSPLIT Multiple Delimiters",
keywords:[
"textsplit multiple delimiter","split by comma and semicolon","multiple delimiter",
"split text","advanced textsplit","excel textsplit","delimiter","parse text","split values","text parsing"
],
formula:'=TEXTSPLIT(A2,{",",";"})',
explanation:"Splits text using multiple delimiters.",
example:"HR,IT;Finance"
},

{
title:"XLOOKUP Multiple Criteria",
keywords:[
"xlookup multiple criteria","lookup two conditions","lookup multiple columns","advanced xlookup",
"employee and department","multi criteria lookup","excel xlookup","multiple conditions","lookup records","two criteria"
],
formula:'=XLOOKUP(1,(A2=A:A)*(B2=B:B),C:C)',
explanation:"Performs lookup using multiple conditions.",
example:"Employee + Department → Salary"
},

{
title:"PMT",
keywords:[
"pmt","loan emi","emi calculation","monthly payment","loan payment",
"mortgage payment","installment","finance","bank loan","emi formula"
],
formula:'=PMT(10%/12,60,-500000)',
explanation:"Calculates the periodic loan payment.",
example:"$500,000 Loan EMI"
},

{
title:"IPMT",
keywords:[
"ipmt","interest payment","loan interest","emi interest","interest installment",
"finance","banking","loan schedule","interest formula","payment interest"
],
formula:'=IPMT(10%/12,1,60,-500000)',
explanation:"Returns the interest portion of a payment.",
example:"Month 1 Interest"
},

{
title:"PPMT",
keywords:[
"ppmt","principal payment","loan principal","emi principal","principal installment",
"finance","banking","loan schedule","principal formula","payment principal"
],
formula:'=PPMT(10%/12,1,60,-500000)',
explanation:"Returns the principal portion of a payment.",
example:"Month 1 Principal"
},

{
title:"FV",
keywords:[
"future value","fv","investment growth","future investment","savings maturity",
"finance","compound interest","investment","future amount","financial planning"
],
formula:'=FV(8%/12,120,-500)',
explanation:"Calculates the future value of an investment.",
example:"10-Year Investment"
},

{
title:"PV",
keywords:[
"present value","pv","loan amount","investment value","discounted value",
"finance","financial planning","current value","loan calculation","present worth"
],
formula:'=PV(8%/12,120,-500)',
explanation:"Returns the present value of an investment or loan.",
example:"Current Investment Value"
},

{
title:"NPV",
keywords:[
"npv","net present value","investment appraisal","cash flow","project valuation",
"finance","discount cash flow","capital budgeting","business valuation","financial analysis"
],
formula:'=NPV(10%,B2:B10)',
explanation:"Calculates the net present value of cash flows.",
example:"Project Profitability"
},

{
title:"XNPV",
keywords:[
"xnpv","net present value dates","cash flow dates","irregular cash flow","investment",
"finance","discount cash flow","financial model","valuation","project analysis"
],
formula:'=XNPV(10%,B2:B10,A2:A10)',
explanation:"Calculates NPV using actual dates.",
example:"Investment Valuation"
},

{
title:"IRR",
keywords:[
"irr","internal rate of return","investment return","project return","financial return",
"finance","capital budgeting","investment analysis","return percentage","irr formula"
],
formula:'=IRR(B2:B10)',
explanation:"Returns the internal rate of return.",
example:"18.5%"
},

{
title:"XIRR",
keywords:[
"xirr","irregular irr","investment return dates","cash flow return","financial analysis",
"finance","actual dates","investment","return percentage","project return"
],
formula:'=XIRR(B2:B10,A2:A10)',
explanation:"Calculates IRR using actual dates.",
example:"22.4%"
},

{
title:"RATE",
keywords:[
"rate","loan interest rate","interest calculation","financial rate","investment rate",
"finance","banking","loan formula","interest percentage","rate function"
],
formula:'=RATE(60,-12000,500000)',
explanation:"Returns the interest rate per period.",
example:"Loan Interest Rate"
},

{
title:"NPER",
keywords:[
"nper","loan tenure","payment periods","loan duration","investment duration",
"finance","financial planning","loan months","number of periods","emi tenure"
],
formula:'=NPER(10%/12,-12000,500000)',
explanation:"Returns the number of payment periods.",
example:"Loan Months"
},

{
title:"CUMIPMT",
keywords:[
"cumipmt","total interest","cumulative interest","loan interest total","bank loan",
"finance","interest paid","loan schedule","financial analysis","interest summary"
],
formula:'=CUMIPMT(10%/12,60,500000,1,12,0)',
explanation:"Returns cumulative interest paid.",
example:"First Year Interest"
},

{
title:"CUMPRINC",
keywords:[
"cumprinc","total principal","principal paid","loan principal total","bank loan",
"finance","loan schedule","financial analysis","principal summary","loan repayment"
],
formula:'=CUMPRINC(10%/12,60,500000,1,12,0)',
explanation:"Returns cumulative principal paid.",
example:"First Year Principal"
},

{
title:"SLN",
keywords:[
"sln","straight line depreciation","asset depreciation","depreciation","fixed asset",
"finance","accounting","asset value","depreciation formula","book value"
],
formula:'=SLN(100000,10000,5)',
explanation:"Calculates straight-line depreciation.",
example:"Annual Depreciation"
},

{
title:"DB",
keywords:[
"db","declining balance depreciation","asset depreciation","fixed asset","finance",
"accounting","depreciation calculation","book value","asset management","db formula"
],
formula:'=DB(100000,10000,5,1)',
explanation:"Returns depreciation using the fixed declining balance method.",
example:"Year 1 Depreciation"
},

{
title:"DDB",
keywords:[
"ddb","double declining balance","accelerated depreciation","asset depreciation","finance",
"accounting","depreciation method","fixed asset","asset value","ddb formula"
],
formula:'=DDB(100000,10000,5,1)',
explanation:"Calculates depreciation using the double declining balance method.",
example:"Year 1 Depreciation"
},

{
title:"SYD",
keywords:[
"syd","sum of years digits","asset depreciation","depreciation method","finance",
"accounting","fixed asset","book value","depreciation formula","asset management"
],
formula:'=SYD(100000,10000,5,1)',
explanation:"Calculates depreciation using the sum-of-years'-digits method.",
example:"Year 1 Depreciation"
},

{
title:"EFFECT",
keywords:[
"effective interest","effective rate","annual interest","compound interest","finance",
"banking","interest calculation","effective annual rate","ear","effective formula"
],
formula:'=EFFECT(12%,12)',
explanation:"Returns the effective annual interest rate.",
example:"12.68%"
},

{
title:"NOMINAL",
keywords:[
"nominal interest","nominal rate","annual nominal rate","finance","banking",
"interest rate","loan rate","nominal formula","interest calculation","financial"
],
formula:'=NOMINAL(12.68%,12)',
explanation:"Returns the nominal annual interest rate.",
example:"12%"
},

{
title:"RRI",
keywords:[
"rri","equivalent interest rate","investment growth","compound rate","finance",
"financial planning","annual growth","investment return","growth rate","rate calculation"
],
formula:'=RRI(5,100000,150000)',
explanation:"Returns the equivalent interest rate for investment growth.",
example:"8.45%"
},

{
title:"PDURATION",
keywords:[
"pduration","investment duration","years to grow","time to investment","financial planning",
"finance","investment period","growth duration","compound growth","duration formula"
],
formula:'=PDURATION(8%,100000,200000)',
explanation:"Returns the number of periods needed for an investment.",
example:"9 Years"
},

{
title:"PRICE",
keywords:[
"bond price","price","security price","bond valuation","finance",
"investment","fixed income","bond formula","market price","financial analysis"
],
formula:'=PRICE(DATE(2025,1,1),DATE(2030,1,1),5%,6%,100,2,0)',
explanation:"Returns the price of a bond per $100 face value.",
example:"Bond Price"
},

{
title:"YIELD",
keywords:[
"bond yield","yield","security yield","investment yield","finance",
"bond return","fixed income","yield formula","market yield","financial analysis"
],
formula:'=YIELD(DATE(2025,1,1),DATE(2030,1,1),5%,95,100,2,0)',
explanation:"Returns the annual yield of a bond.",
example:"6.18%"
},

{
title:"DURATION",
keywords:[
"duration","bond duration","macaulay duration","interest risk","finance",
"bond analysis","fixed income","duration formula","investment","risk analysis"
],
formula:'=DURATION(DATE(2025,1,1),DATE(2030,1,1),5%,6%,2,0)',
explanation:"Returns the Macaulay duration of a bond.",
example:"4.28 Years"
},

{
title:"MDURATION",
keywords:[
"modified duration","mduration","bond sensitivity","interest rate risk","finance",
"bond duration","fixed income","risk management","investment","bond analysis"
],
formula:'=MDURATION(DATE(2025,1,1),DATE(2030,1,1),5%,6%,2,0)',
explanation:"Returns the modified duration of a bond.",
example:"4.05 Years"
},

{
title:"NETWORKDAYS",
keywords:[
"working days","networkdays","business days","exclude weekends","working day calculation",
"office days","attendance","hr","employee working days","calculate weekdays"
],
formula:'=NETWORKDAYS(A2,B2)',
explanation:"Returns the number of working days between two dates.",
example:"01-Jan-2026 to 31-Jan-2026 → 22 Days"
},

{
title:"NETWORKDAYS.INTL",
keywords:[
"networkdays intl","custom weekend","working days custom","business days","custom holidays",
"attendance","hr","office calendar","weekend calculation","intl working days"
],
formula:'=NETWORKDAYS.INTL(A2,B2,1,Holidays)',
explanation:"Calculates working days using custom weekends and holidays.",
example:"Working Days Excluding Holidays"
},

{
title:"WORKDAY",
keywords:[
"workday","next working day","future workday","business day","office date",
"hr","attendance","working date","excel workday","calculate workday"
],
formula:'=WORKDAY(A2,10)',
explanation:"Returns a future working date.",
example:"10 Working Days Later"
},

{
title:"WORKDAY.INTL",
keywords:[
"workday intl","custom workday","future working day","holiday calendar","custom weekend",
"attendance","office","hr","business calendar","intl workday"
],
formula:'=WORKDAY.INTL(A2,10,1,Holidays)',
explanation:"Returns a future working day using custom weekends.",
example:"Next Working Date"
},

{
title:"DATEDIF",
keywords:[
"datedif","age calculation","difference between dates","service years","experience",
"employee experience","date difference","years months days","hr","age formula"
],
formula:'=DATEDIF(A2,B2,"Y")',
explanation:"Calculates difference between two dates.",
example:"Joining Date → Experience"
},

{
title:"EOMONTH",
keywords:[
"end of month","eomonth","month ending","last day month","month close",
"finance","salary month","accounting","month end","reporting"
],
formula:'=EOMONTH(A2,0)',
explanation:"Returns the last day of the month.",
example:"31-Jan-2026"
},

{
title:"EDATE",
keywords:[
"edate","add months","future month","previous month","salary increment date",
"contract renewal","loan date","month calculation","date add","excel edate"
],
formula:'=EDATE(A2,6)',
explanation:"Adds or subtracts months from a date.",
example:"6 Months Later"
},

{
title:"YEARFRAC",
keywords:[
"yearfrac","fraction year","service period","tenure","experience calculation",
"finance","employee service","year difference","decimal years","hr"
],
formula:'=YEARFRAC(A2,B2)',
explanation:"Returns the fraction of a year between two dates.",
example:"3.75 Years"
},

{
title:"TEXT DATE",
keywords:[
"format date","date format","convert date","text date","display date",
"dd mmm yyyy","custom date","excel date format","date text","formatted date"
],
formula:'=TEXT(A2,"dd-mmm-yyyy")',
explanation:"Formats a date as text.",
example:"15-Jul-2026"
},

{
title:"WEEKNUM",
keywords:[
"week number","weeknum","calendar week","week calculation","current week",
"week of year","excel week","reporting week","week formula","weekly report"
],
formula:'=WEEKNUM(A2)',
explanation:"Returns the week number of the year.",
example:"Week 32"
},

{
title:"ISOWEEKNUM",
keywords:[
"iso week","iso week number","international week","iso calendar","weeknum iso",
"reporting","calendar","week calculation","excel iso","week"
],
formula:'=ISOWEEKNUM(A2)',
explanation:"Returns the ISO week number.",
example:"ISO Week 31"
},

{
title:"FILTER Active Employees",
keywords:[
"active employees","filter active","employee status","active staff","current employees",
"hr dashboard","employee list","attendance","working employees","staff filter"
],
formula:'=FILTER(A2:F100,F2:F100="Active")',
explanation:"Returns only active employees.",
example:"Current Active Staff"
},

{
title:"FILTER Department",
keywords:[
"department filter","filter department","hr department","employee department","department wise",
"filter hr","department records","staff by department","employee list","department data"
],
formula:'=FILTER(A2:F100,C2:C100="Finance")',
explanation:"Returns employees from a specific department.",
example:"Finance Employees"
},

{
title:"UNIQUE Departments",
keywords:[
"unique departments","department list","distinct department","remove duplicates department",
"unique values","department names","hr report","department summary","distinct list","unique"
],
formula:'=UNIQUE(C2:C100)',
explanation:"Returns unique department names.",
example:"HR, Finance, IT"
},

{
title:"SORT Salary Descending",
keywords:[
"highest salary","salary descending","sort salary","top salary","salary ranking",
"employee salary","sort employees","salary report","highest paid","salary sort"
],
formula:'=SORTBY(A2:F100,E2:E100,-1)',
explanation:"Sorts employees by salary in descending order.",
example:"Highest Paid Employee First"
},

{
title:"TOP 10 Employees",
keywords:[
"top 10 employees","top performers","highest salary employees","best employees","top records",
"employee ranking","performance report","top staff","excel top 10","dashboard"
],
formula:'=TAKE(SORTBY(A2:F100,E2:E100,-1),10)',
explanation:"Returns the top 10 employees based on salary.",
example:"Top 10 Salary List"
},

{
title:"Attendance Percentage",
keywords:[
"attendance percentage","attendance %","employee attendance","present percentage","attendance report",
"hr dashboard","attendance calculation","employee performance","attendance formula","present ratio"
],
formula:'=COUNTIF(B2:AF2,"P")/COUNTA(B2:AF2)',
explanation:"Calculates employee attendance percentage.",
example:"96%"
},

{
title:"Leave Balance",
keywords:[
"leave balance","remaining leave","employee leave","annual leave","leave tracker",
"hr leave","leave calculation","vacation balance","leave management","available leave"
],
formula:'=AllocatedLeave-TakenLeave',
explanation:"Calculates remaining leave balance.",
example:"24 - 6 = 18 Days"
},

{
title:"Overtime Hours",
keywords:[
"overtime","ot hours","extra hours","working overtime","employee overtime",
"attendance","timesheet","working hours","ot calculation","extra work"
],
formula:'=(OutTime-InTime)*24-8',
explanation:"Calculates overtime hours after an 8-hour shift.",
example:"2.5 Hours OT"
},

{
title:"Experience in Years",
keywords:[
"employee experience","experience years","service years","joining date","employee tenure",
"hr report","years of service","work experience","experience calculation","employee age"
],
formula:'=DATEDIF(A2,TODAY(),"Y")',
explanation:"Calculates employee experience in years.",
example:"8 Years"
},

{
title:"Retirement Date",
keywords:[
"retirement date","employee retirement","retire age","superannuation","retirement calculation",
"hr planning","employee exit","retirement year","age 60","service planning"
],
formula:'=EDATE(DateOfBirth,60*12)',
explanation:"Calculates retirement date assuming retirement at age 60.",
example:"15-Aug-2048"
},

{
title:"Payroll Month",
keywords:[
"payroll month","salary month","month name","pay period","salary processing",
"payroll report","month text","salary cycle","pay month","finance"
],
formula:'=TEXT(A2,"mmmm-yyyy")',
explanation:"Returns payroll month.",
example:"July-2026"
},

{
title:"Bonus Calculation",
keywords:[
"bonus","employee bonus","annual bonus","bonus amount","performance bonus",
"salary bonus","payroll","hr bonus","bonus formula","incentive"
],
formula:'=Salary*10%',
explanation:"Calculates a 10% bonus.",
example:"$5,000 → $500"
},

{
title:"Net Salary",
keywords:[
"net salary","take home salary","salary after deduction","payroll","employee salary",
"net pay","salary calculation","monthly salary","ctc","salary sheet"
],
formula:'=GrossSalary-Deductions',
explanation:"Calculates net salary after deductions.",
example:"$5,000 - $500 = $4,500"
},

{
title:"Payroll Status",
keywords:[
"payroll status","salary processed","payment status","paid pending","salary release",
"payroll dashboard","salary report","payment tracking","employee payment","pay status"
],
formula:'=IF(PaymentDate<>"","Paid","Pending")',
explanation:"Shows payroll processing status.",
example:"Paid"
},

{
title:"Treasury Position",
keywords:[
"treasury position","cash position","daily treasury","cash balance","liquidity",
"treasury report","bank treasury","opening balance","closing balance","cash management"
],
formula:'=SUM(OpeningBalance:ClosingBalance)',
explanation:"Calculates total treasury cash position.",
example:"Daily Treasury Balance"
},

{
title:"FX Gain Loss",
keywords:[
"fx gain loss","foreign exchange","forex profit","currency gain","exchange loss",
"fx report","currency fluctuation","treasury fx","forex calculation","gain loss"
],
formula:'=(SellingRate-BuyingRate)*Amount',
explanation:"Calculates foreign exchange gain or loss.",
example:"USD Transaction Profit"
},

{
title:"Dealer Limit Check",
keywords:[
"dealer limit","trading limit","treasury dealer","limit exceeded","dealer exposure",
"risk limit","trading control","limit monitoring","dealer check","compliance"
],
formula:'=IF(A2>Limit,"Limit Exceeded","Within Limit")',
explanation:"Checks whether dealer trading limit is exceeded.",
example:"Limit Status"
},

{
title:"SLA Status",
keywords:[
"sla","service level","sla status","within sla","breach","turnaround time",
"tat","service compliance","support sla","ticket sla"
],
formula:'=IF(ResolvedTime<=TargetTime,"Within SLA","SLA Breach")',
explanation:"Checks SLA compliance.",
example:"Within SLA"
},

{
title:"TAT Hours",
keywords:[
"tat","turnaround time","processing time","resolution hours","ticket time",
"business hours","case duration","response time","service time","elapsed hours"
],
formula:'=(EndTime-StartTime)*24',
explanation:"Calculates turnaround time in hours.",
example:"5.25 Hours"
},

{
title:"Ageing Bucket",
keywords:[
"ageing","aging","overdue","pending ageing","bucket","days pending",
"invoice ageing","ticket ageing","age analysis","delay bucket"
],
formula:'=IFS(A2<=30,"0-30",A2<=60,"31-60",A2<=90,"61-90",TRUE,"90+")',
explanation:"Categorizes records into ageing buckets.",
example:"31-60 Days"
},

{
title:"Compliance Status",
keywords:[
"compliance","compliant","non compliant","audit status","control status",
"regulatory","policy compliance","risk compliance","icofr","sox"
],
formula:'=IF(A2="Yes","Compliant","Non-Compliant")',
explanation:"Returns compliance status.",
example:"Compliant"
},

{
title:"Mandatory Leave Check",
keywords:[
"mandatory leave","planned mandatory leave","pml","leave compliance","annual leave",
"treasury leave","mandatory vacation","leave policy","rbi leave","leave check"
],
formula:'=IF(LeaveDays>=10,"Compliant","Pending")',
explanation:"Checks mandatory leave compliance.",
example:"Compliant"
},

{
title:"Job Rotation Status",
keywords:[
"job rotation","rotation compliance","employee rotation","dealer rotation",
"rbi rotation","role rotation","staff rotation","rotation due","rotation tracker","rotation status"
],
formula:'=IF(TODAY()-LastRotation>=365,"Due","Completed")',
explanation:"Checks annual job rotation status.",
example:"Due"
},

{
title:"Duplicate Check",
keywords:[
"duplicate","duplicate records","duplicate employee","duplicate data",
"find duplicates","remove duplicate","duplicate id","duplicate values","excel duplicate","unique check"
],
formula:'=IF(COUNTIF(A:A,A2)>1,"Duplicate","Unique")',
explanation:"Identifies duplicate values.",
example:"Duplicate"
},

{
title:"MIS Completion %",
keywords:[
"mis completion","dashboard progress","report completion","tracker progress",
"activity completion","status percentage","project completion","progress","kpi completion","mis tracker"
],
formula:'=CompletedTasks/TotalTasks',
explanation:"Calculates completion percentage.",
example:"92%"
},

{
title:"Risk Rating",
keywords:[
"risk rating","risk level","high risk","medium risk","low risk",
"risk assessment","control rating","operational risk","risk score","risk category"
],
formula:'=IFS(A2>=80,"High",A2>=50,"Medium",TRUE,"Low")',
explanation:"Assigns a risk category based on score.",
example:"Medium Risk"
},

{
title:"Control Effectiveness",
keywords:[
"control effectiveness","control rating","effective control","icofr testing",
"internal control","audit control","control review","risk control","testing result","control status"
],
formula:'=IF(TestResult="Pass","Effective","Needs Improvement")',
explanation:"Evaluates control effectiveness.",
example:"Effective"
},

{
title:"Dashboard KPI",
keywords:[
"kpi","dashboard kpi","performance indicator","business dashboard","scorecard",
"management dashboard","executive dashboard","performance score","kpi status","dashboard metric"
],
formula:'=Actual/Target',
explanation:"Calculates KPI achievement percentage.",
example:"96%"
},

{
title:"Power Query Refresh Date",
keywords:[
"power query refresh","last refresh","data refresh","query refresh",
"etl refresh","refresh time","power query date","data updated","refresh status","pq refresh"
],
formula:'=NOW()',
explanation:"Displays the last data refresh date and time.",
example:"26-Jul-2026 10:30 AM"
},

{
title:"Variance",
keywords:[
"variance","actual vs target","budget variance","difference","variance analysis",
"financial variance","sales variance","performance variance","actual budget","gap analysis"
],
formula:'=Actual-Target',
explanation:"Calculates variance between actual and target.",
example:"$1,250"
},

{
title:"Variance %",
keywords:[
"variance percent","variance %","percentage variance","budget variance percentage",
"performance variance","actual vs target %","gap percentage","variance ratio","growth variance","financial analysis"
],
formula:'=(Actual-Target)/Target',
explanation:"Calculates percentage variance.",
example:"12%"
},

{
title:"Achievement %",
keywords:[
"achievement","achievement percentage","target achieved","goal completion",
"performance percentage","kpi achievement","sales achievement","employee target","target completion","achievement rate"
],
formula:'=Actual/Target',
explanation:"Calculates achievement percentage.",
example:"108%"
},

{
title:"Trend",
keywords:[
"trend","growth trend","performance trend","increase decrease","monthly trend",
"business trend","sales trend","kpi trend","trend analysis","performance direction"
],
formula:'=IF(CurrentMonth>PreviousMonth,"Up","Down")',
explanation:"Shows business trend direction.",
example:"Up"
},

{
title:"Status Indicator",
keywords:[
"status","traffic light","red amber green","rag status","indicator",
"performance status","dashboard color","health status","progress status","overall status"
],
formula:'=IFS(A2>=90,"🟢 Green",A2>=75,"🟡 Amber",TRUE,"🔴 Red")',
explanation:"Displays dashboard status indicator.",
example:"🟢 Green"
},

{
title:"Power Query Append Check",
keywords:[
"append","append files","power query append","merge excel","combine files",
"multiple excel","folder import","append csv","etl append","combine workbook"
],
formula:'=COUNTA(AppendedData)',
explanation:"Returns total records after appending multiple files.",
example:"15,280 Records"
},

{
title:"Power Query Duplicate Removal",
keywords:[
"remove duplicates","power query duplicate","clean data","duplicate removal",
"data cleaning","etl","unique records","remove repeated","power query clean","duplicate check"
],
formula:'=ROWS(UNIQUE(A2:A1000))',
explanation:"Counts unique records after removing duplicates.",
example:"9,850 Unique Records"
},

{
title:"Power BI Refresh Status",
keywords:[
"power bi refresh","dataset refresh","refresh failed","refresh success",
"power bi dataset","dashboard refresh","bi refresh","refresh monitoring","power bi status","dataset status"
],
formula:'=IF(A2="Success","Updated","Failed")',
explanation:"Shows Power BI dataset refresh status.",
example:"Updated"
},

{
title:"Audit Observation Status",
keywords:[
"audit observation","audit issue","open observation","closed observation",
"internal audit","external audit","audit tracker","observation status","audit finding","audit report"
],
formula:'=IF(A2="Closed","Completed","Open")',
explanation:"Displays audit observation status.",
example:"Completed"
},

{
title:"Overall Dashboard Score",
keywords:[
"overall score","dashboard score","performance score","management score",
"overall performance","kpi score","business score","executive score","dashboard rating","score"
],
formula:'=AVERAGE(B2:F2)',
explanation:"Calculates the overall dashboard performance score.",
example:"94.6%"
},

{
title:"Dynamic Drop Down",
keywords:[
"dynamic dropdown","dependent dropdown","dropdown list","validation list",
"dynamic list","excel dropdown","data validation","search dropdown","list selection","dropdown"
],
formula:'=SORT(UNIQUE(A2:A100))',
explanation:"Creates a dynamic dropdown source list.",
example:"Department Dropdown"
},

{
title:"Dynamic Search",
keywords:[
"dynamic search","search records","excel search","lookup search","filter search",
"search employee","live search","find record","search data","textbox search"
],
formula:'=FILTER(A2:F100,ISNUMBER(SEARCH(H2,A2:A100)),"No Results")',
explanation:"Searches records dynamically.",
example:"Search Employee Name"
},

{
title:"Top 5 Records",
keywords:[
"top 5","highest 5","top records","best records","top values",
"highest values","dashboard top","ranking","top employees","top sales"
],
formula:'=TAKE(SORTBY(A2:F100,E2:E100,-1),5)',
explanation:"Returns the top 5 records.",
example:"Top 5 Sales Employees"
},

{
title:"Bottom 5 Records",
keywords:[
"bottom 5","lowest 5","lowest records","bottom values","smallest values",
"lowest sales","worst performers","bottom employees","lowest ranking","bottom records"
],
formula:'=TAKE(SORTBY(A2:F100,E2:E100,1),5)',
explanation:"Returns the bottom 5 records.",
example:"Lowest Sales Employees"
},

{
title:"Running Total",
keywords:[
"running total","cumulative sum","running balance","progressive total",
"accumulated total","rolling sum","balance running","sales running total","running amount","cumulative"
],
formula:'=SUM($B$2:B2)',
explanation:"Calculates a running total.",
example:"Daily Sales Running Total"
},

{
title:"Percent of Total",
keywords:[
"percent of total","percentage contribution","share percentage","ratio",
"contribution","total percentage","dashboard percentage","sales contribution","part of total","% total"
],
formula:'=B2/SUM(B:B)',
explanation:"Calculates contribution to the total.",
example:"8.25%"
},

{
title:"Monthly Growth %",
keywords:[
"monthly growth","mom growth","growth percentage","month over month",
"sales growth","business growth","increase percentage","growth analysis","monthly comparison","trend growth"
],
formula:'=(B3-B2)/B2',
explanation:"Calculates month-over-month growth.",
example:"12.5%"
},

{
title:"Year over Year Growth",
keywords:[
"yoy","year over year","annual growth","yearly growth","growth comparison",
"sales growth yearly","business growth","compare years","yearly trend","yoy growth"
],
formula:'=(B2-A2)/A2',
explanation:"Calculates year-over-year growth.",
example:"18%"
},

{
title:"Dynamic Chart Range",
keywords:[
"dynamic chart","chart range","auto chart","expand chart","dynamic graph",
"chart source","dashboard chart","excel chart","live chart","graph data"
],
formula:'=OFFSET(A1,0,0,COUNTA(A:A),2)',
explanation:"Creates a dynamic chart range.",
example:"Auto Updating Chart"
},

{
title:"Last Non Blank Value",
keywords:[
"last value","last non blank","latest value","recent record","last entry",
"last data","recent data","final value","last filled cell","latest record"
],
formula:'=LOOKUP(2,1/(A:A<>""),A:A)',
explanation:"Returns the last non-empty value.",
example:"Latest Transaction"
},

{
title:"First Non Blank Value",
keywords:[
"first value","first non blank","first record","first entry","initial value",
"first data","starting value","first filled cell","beginning record","first transaction"
],
formula:'=INDEX(A:A,MATCH(TRUE,A:A<>"",0))',
explanation:"Returns the first non-empty value.",
example:"First Employee"
},

{
title:"Dynamic Serial Number",
keywords:[
"serial number","auto serial","dynamic numbering","sequence rows","row numbering",
"automatic numbering","excel serial","record number","index number","sl no"
],
formula:'=SEQUENCE(COUNTA(A2:A100))',
explanation:"Generates automatic serial numbers.",
example:"1,2,3,4..."
},

{
title:"Duplicate Highlighter",
keywords:[
"highlight duplicate","duplicate format","duplicate values","conditional duplicate",
"duplicate check","find duplicate","duplicate color","repeat values","duplicate data","duplicate records"
],
formula:'=COUNTIF($A:$A,A2)>1',
explanation:"Highlights duplicate records using Conditional Formatting.",
example:"Duplicate Employee ID"
},

{
title:"Weekend Check",
keywords:[
"weekend","holiday check","saturday sunday","weekend formula","day check",
"working day","office day","calendar","attendance","weekend validation"
],
formula:'=IF(WEEKDAY(A2,2)>5,"Weekend","Weekday")',
explanation:"Checks whether a date falls on a weekend.",
example:"Saturday → Weekend"
},

{
title:"Quarter",
keywords:[
"quarter","financial quarter","q1 q2 q3 q4","quarter number","fiscal quarter",
"quarter report","quarterly","business quarter","quarter calculation","financial reporting"
],
formula:'="Q"&ROUNDUP(MONTH(A2)/3,0)',
explanation:"Returns the financial quarter.",
example:"Q3"
},

{
title:"Fiscal Year",
keywords:[
"fiscal year","financial year","fy","accounting year","financial reporting",
"year ending","india financial year","business year","fy calculation","financial calendar"
],
formula:'=IF(MONTH(A2)>=4,YEAR(A2)&"-"&YEAR(A2)+1,YEAR(A2)-1&"-"&YEAR(A2))',
explanation:"Returns the fiscal year (Apr-Mar).",
example:"2026-2027"
},

{
title:"Email Validation",
keywords:[
"email validation","valid email","email check","email format","verify email",
"email address","invalid email","email test","email formula","mail validation"
],
formula:'=ISNUMBER(SEARCH("@",A2))',
explanation:"Checks if an email appears valid.",
example:"TRUE"
},

{
title:"Phone Validation",
keywords:[
"phone validation","mobile validation","mobile number","phone check","10 digit",
"telephone","contact number","mobile format","phone verify","number validation"
],
formula:'=AND(LEN(A2)=10,ISNUMBER(A2))',
explanation:"Validates a 10-digit phone number.",
example:"TRUE"
},

{
title:"Age Calculator",
keywords:[
"age","age calculator","calculate age","employee age","birth date",
"date of birth","dob","years old","age formula","birthday"
],
formula:'=DATEDIF(A2,TODAY(),"Y")',
explanation:"Calculates age in years.",
example:"32 Years"
},

{
title:"Working Hours",
keywords:[
"working hours","office hours","shift hours","attendance hours","login logout",
"time difference","work duration","employee hours","timesheet","office timing"
],
formula:'=(B2-A2)*24',
explanation:"Calculates total working hours.",
example:"8.5 Hours"
},

{
title:"Late Coming",
keywords:[
"late coming","late attendance","employee late","shift delay","office late",
"attendance delay","late check","arrival time","late mark","hr attendance"
],
formula:'=IF(B2>TIME(9,15,0),"Late","On Time")',
explanation:"Checks whether an employee is late.",
example:"Late"
},

{
title:"Early Leaving",
keywords:[
"early leaving","early exit","left early","office exit","attendance",
"employee exit","shift end","early logout","leaving time","hr"
],
formula:'=IF(C2<TIME(18,0,0),"Early Exit","Completed Shift")',
explanation:"Checks if an employee left before shift end.",
example:"Completed Shift"
},

{
title:"Power Query Refresh Status",
keywords:[
"power query refresh","refresh status","query updated","etl refresh",
"data refresh","refresh check","last refresh","power query","update status","refresh time"
],
formula:'=NOW()',
explanation:"Displays the latest Power Query refresh timestamp.",
example:"27-Jul-2026 09:15 AM"
},

{
title:"Dashboard Last Updated",
keywords:[
"dashboard updated","last updated","report refresh","dashboard refresh",
"last refresh","report date","dashboard timestamp","latest update","update time","refresh"
],
formula:'=NOW()',
explanation:"Displays dashboard last updated date and time.",
example:"27-Jul-2026 09:15 AM"
},

{
title:"AI Formula Suggestion",
keywords:[
"ai formula","suggest formula","formula recommendation","excel ai","generate formula",
"best formula","formula finder","formula generator","ai assistant","excel helper"
],
formula:'="Describe your problem and AI will suggest the best Excel formula."',
explanation:"Placeholder for your future AI engine.",
example:"'Calculate bonus based on performance' → AI Recommendation"
}

];

// =======================================
// Search Formula
// =======================================

generateBtn.addEventListener("click", function(){

const search = problem.value.trim().toLowerCase();

if(search==""){

alert("Please enter your Excel problem.");

return;

}

const results = FORMULAS.filter(item=>{

return item.keywords.some(keyword=>

keyword.toLowerCase().includes(search) ||

search.includes(keyword.toLowerCase())

);

});

if(results.length==0){

formula.innerHTML="<h4>🤖 No Formula Found</h4>";

explanation.innerHTML="Try another keyword.";

example.innerHTML="Example: XLOOKUP, SUMIFS, IF, FILTER";

return;

}

formula.innerHTML="";

explanation.innerHTML="";

example.innerHTML="";

results.forEach(item=>{

formula.innerHTML += `
<div style="border:1px solid #ddd;padding:15px;margin-bottom:15px;border-radius:10px;background:#fff;">
<h4 style="color:#107C41;">${item.title}</h4>
<pre>${item.formula}</pre>
</div>
`;

explanation.innerHTML += `
<p><b>${item.title}</b><br>${item.explanation}</p>
`;

example.innerHTML += `
<p><b>${item.title}</b><br>${item.example}</p>
`;

});

});




// =======================================
// Copy
// =======================================

copyBtn.addEventListener("click",()=>{

navigator.clipboard.writeText(formula.innerText);

copyBtn.innerText="✅ Copied";

setTimeout(()=>{

copyBtn.innerText="📋 Copy Formula";

},1500);

});

