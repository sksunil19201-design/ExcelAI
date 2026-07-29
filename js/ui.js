/*
=========================================
 Excel AI
 ui.js
 Version : 1.0
=========================================
*/

(function(){

// ===============================
// Toast Message
// ===============================

function showToast(message){

    let toast=document.getElementById("toast");

    if(!toast){

        toast=document.createElement("div");

        toast.id="toast";

        document.body.appendChild(toast);

    }

    toast.innerHTML=message;

    toast.classList.add("show");

    setTimeout(()=>{

        toast.classList.remove("show");

    },2000);

}

// ===============================
// Copy Formula
// ===============================

const copyBtn=document.getElementById("copyFormula");

if(copyBtn){

copyBtn.addEventListener("click",()=>{

const formula=document.getElementById("formulaBox").innerText;

navigator.clipboard.writeText(formula);

showToast("✅ Formula Copied");

});

}

// ===============================
// Copy Complete Answer
// ===============================

const copyAll=document.getElementById("copyAll");

if(copyAll){

copyAll.addEventListener("click",()=>{

let text="";

text+="Formula\n";

text+=document.getElementById("formulaBox").innerText+"\n\n";

text+="Explanation\n";

text+=document.getElementById("explanation").innerText+"\n\n";

text+="Business Use Case\n";

text+=document.getElementById("businessCase").innerText+"\n\n";

text+="Alternative Formula\n";

text+=document.getElementById("alternativeFormula").innerText;

navigator.clipboard.writeText(text);

showToast("✅ Answer Copied");

});

}

// ===============================
// New Chat
// ===============================

const newChat=document.getElementById("newChatBtn");

if(newChat){

newChat.addEventListener("click",()=>{

document.getElementById("prompt").value="";

document.getElementById("formulaBox").innerText="Waiting for AI...";

document.getElementById("explanation").innerText="Waiting for response...";

document.getElementById("businessCase").innerText="Waiting...";

document.getElementById("alternativeFormula").innerText="-";

document.getElementById("tipsList").innerHTML="";

document.getElementById("errorList").innerHTML="";

document.getElementById("breakdown").innerHTML="Waiting...";

document.getElementById("outputValue").innerHTML="0";

document.getElementById("tableContainer").innerHTML="";

showToast("🆕 New Chat Started");

});

}

// ===============================
// Dark Mode
// ===============================

const theme=document.getElementById("themeBtn");

if(theme){

theme.addEventListener("click",()=>{

document.body.classList.toggle("dark");

showToast("🌙 Theme Changed");

});

}

window.showToast=showToast;

console.log("✅ ui.js Loaded");

})();
