import {userMessage} from './mots.js';
import {botReply} from './reponses.js';
import {alternative} from './alternatives.js';
import {commonWords} from './commonWords.js';


const button =
document.getElementById("button");
button. addEventListener('click', ()=>{sendMessage();});

 




// Ensure the number of responses in botReply matches the number of message groups in userMessage
if (botReply.length !== userMessage.length) {
  console.log("The number of message groups in botReply must match the number of message groups in userMessage.");
  console.log(userMessage.length)
  console.log(botReply.length);
}


let lastUserDiv;

function sendMessage() {
  const inputField = document.getElementById("input");
  const input = inputField.value.trim();
  if (input) {
    output(input);
    inputField.value = "";
  }
}

//Calculation 
function findResponse(input) {

  //checks direct messages
  let text = cleanInput(input);
  for (const messageGroup of userMessage) {
    const matchingMessage = messageGroup.find(message => text.includes(message.toLowerCase()));
    if (matchingMessage) {
      return getRandomResponseFromCategory(userMessage.indexOf(messageGroup));
    }
  }
  const directMatch = findDirectMatch(text);
  if (directMatch) {
    return directMatch;
  }

  //check similarity
  const SIMILARITY_THRESHOLD = 0.6;
  for (let i = 0; i < userMessage.length; i++) {
    const messageGroup = userMessage[i];
    for (const message of messageGroup) {
      const similarityScore = calculateSimilarity(text, message);
      if (similarityScore >= SIMILARITY_THRESHOLD) {
        return getRandomResponseFromCategory(i);
      }
    }
  }

  //check each word
  const words = text.split(' ');
  for (const word of words) {
    if (isCommonWord(word)) {
      continue;
    }
    const categoryIndex = findCategoryIndex(word);
    if (categoryIndex !== -1) {
      return getRandomResponseFromCategory(categoryIndex);
    }
  }

  return getRandomAlternativeResponse();
}

function calculateSimilarity(str1, str2) {
  const len1 = str1.length;
  const len2 = str2.length;
  const short = len1 < len2 ? len1 : len2;
  const long = len1 > len2 ? len1 : len2;
  let matchCount = 0;

  for (let i = 0; i < short; i++) {
    if (str1[i] === str2[i]) {
      matchCount++;
    }
  }
  const similarity = matchCount / long;
  return similarity;
}

function cleanInput(input) {
  return input
    .toLowerCase()
    .replace(/[^\w\s'"]/gi, "")
    .replace(/[\W_]/g, " ")
    .replace(/ a /g, " ")
    .replace(/i feel /g, "")
    .replace(/whats/g, "what is")
    .replace(/please /g, "")
    .replace(/ please/g, "")
    .trim();
}

function findDirectMatch(text) {
  for (const messageGroup of userMessage) {
    const matchingMessage = messageGroup.find(message => text.includes(message.toLowerCase()));
    if (matchingMessage) {
      return getRandomResponseFromCategory(userMessage.indexOf(messageGroup));
    }
  }
  return null;
}

function isCommonWord(word) {
  
  return commonWords.includes(word.toLowerCase());
}

function findCategoryIndex(word) {
  return botReply.findIndex(category => category.includes(word));
}

function getRandomResponseFromCategory(categoryIndex) {
  const responses = botReply[categoryIndex];
  return responses[Math.floor(Math.random() * responses.length)];
}

function getRandomAlternativeResponse() {
  return alternative[Math.floor(Math.random() * alternative.length)];
}

document.addEventListener("DOMContentLoaded", () => {
  const inputField = document.getElementById("input");
  inputField.addEventListener("keydown", function (e) {
    if (e.code === "Enter") {
      const input = inputField.value.trim();
      if (input) {
        output(input);
        inputField.value = "";
      }
    }
  });
});

const usnexp = /\b4AI\d{2}AI\d{3}\b/i;

function output(input) {
  if (input.toLowerCase().includes("university") || usnexp.test(input.toLowerCase()) || input.toLowerCase().includes("usn")) {
    userMessageSend(input);
    const match = usnexp.test(input);
    if (match) {
      const userUSN = input.match(usnexp)[0];
      readStudentDatafromxl(userUSN);
    } else {
      sendBotMessage("Enter USN 📝");
      // Call the userMessageWait function
      userMessageWait().then((userResponse) => {
        userResponse = cleanInput(userResponse).replace(/\s/g, "");
        const match = usnexp.test(userResponse);
        if (match) {
          readStudentDatafromxl(userResponse);
        }
      });
    }
  } else {
    // If the input does not contain the USN pattern, proceed with generating a response
    let product = findResponse(input);
    addChat(input, product);
  }
}

function readStudentDatafromxl(userUSN) {
  const fileName = "students.xlsx"; // Specify the name of the Excel file

  // Fetch the Excel file using the Fetch API
  fetch(fileName)
    .then((response) => response.arrayBuffer())
    .then((data) => {
      const workbook = XLSX.read(data, { type: "array" });

      // Assuming the sheet name is "student"
      const sheetName = workbook.SheetNames[0];
      const worksheet = workbook.Sheets[sheetName];

      // Find the row with the matching USN (case-insensitive)
      const rows = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
      const headerRow = rows[0];
      let found = false;

      for (let i = 1; i < rows.length; i++) {
        const row = rows[i];
        const usnIndex = headerRow.indexOf("USN");
        const usnFromSheet = row[usnIndex].toLowerCase(); // Convert USN from the sheet to lowercase
        if (usnFromSheet === userUSN.toLowerCase()) { // Convert user input to lowercase for comparison
          // Assuming the columns are in this order: Name, USN, Email, Phone, Club, Year
          const nameIndex = headerRow.indexOf("Name");
          const emailIndex = headerRow.indexOf("Email");
          const phoneIndex = headerRow.indexOf("Phone");
          const clubIndex = headerRow.indexOf("Club");
          const yearIndex = headerRow.indexOf("Year");
          const name = row[nameIndex];
          const email = row[emailIndex];
          const phone = row[phoneIndex];
          const club = row[clubIndex];
          const year = row[yearIndex];
          const result = `Name: ${name}<br>USN: ${userUSN.toUpperCase()}<br>Email: ${email}<br>Phone: ${phone}<br>Club: ${club}<br>Year: ${year}`;
          // Call sendBotMessage with the result
          sendBotMessage(result, `Searching for USN ${userUSN.toUpperCase()} ...`);
          found = true;
          break;
        }
      }
      if (!found) {
        sendBotMessage("Student not found!");
      }
    })
    .catch((error) => {
      console.error("Error loading the Excel file:", error);
    });
}

const inputField = document.getElementById("input");
const quickMessageButtons = document.querySelectorAll(".quickmessage");
const mainDiv = document.getElementById("chatBox");
const botDiv = document.createElement("div");

function sendBotMessage(message, typingText = "Typing...") {
  botDiv.id = "bot";
  botDiv.classList.add("message");
  botDiv.innerHTML = `<span id="bot-response" class="typing-animation">${typingText}</span>`;
  mainDiv.appendChild(botDiv);
  var scroll = document.getElementById("chatBox");
  scroll.scrollTop = scroll.scrollHeight;
  inputField.disabled = true;
  quickMessageButtons.forEach(button => {
    button.disabled = true;
  });

  setTimeout(() => {
    const botResponse = botDiv.querySelector("#bot-response");
    botResponse.classList.remove("typing-animation");
    botDiv.innerHTML = `<span id="bot-response">${message}</span>`;
    scroll.scrollTop = scroll.scrollHeight;
    inputField.disabled = false;
    quickMessageButtons.forEach(button => {
      button.disabled = false;
    });

    inputField.focus();
  }, 2000);
}

function userMessageSend(input) {
  // Create a new user message div
  const newUserDiv = document.createElement("div");
  newUserDiv.id = "user";
  newUserDiv.classList.add("message", "typing-animation");
  newUserDiv.style.animationDuration = "1s";
  newUserDiv.innerHTML = `<span id="user-response">${input}</span>`;
  mainDiv.appendChild(newUserDiv);

  var scroll = document.getElementById("chatBox");
  scroll.scrollTop = scroll.scrollHeight;

  // Remove "typing-animation" class after 500ms
  setTimeout(() => {
    newUserDiv.classList.remove("typing-animation");
  }, 500);
}

function userMessageWait() {
  return new Promise((resolve) => {
    const inputField = document.getElementById("input");
    let userResponse = "";

    function handleUserResponse(e) {
      if (e.code === "Enter" || e.target.classList.contains("sendmessage")) {
        const trimmedResponse = userResponse.trim();
        resolve(trimmedResponse);
      } else {
        userResponse = inputField.value;
      }
    }

    function handleInputBlur() {
      resolve(userResponse.trim());
    }

    inputField.addEventListener("keydown", handleUserResponse);
    inputField.addEventListener("input", handleUserResponse);
    inputField.addEventListener("blur", handleInputBlur);
  });
}


function addChat(input, product) {
  // Create a new user message div
  const newUserDiv = document.createElement("div");
  newUserDiv.id = "user";
  newUserDiv.classList.add("message", "typing-animation");
  newUserDiv.style.animationDuration = "0.5s";
  newUserDiv.innerHTML = `<span id="user-response">${input}</span>`;
  mainDiv.appendChild(newUserDiv);

  const botDiv = document.createElement("div");
  botDiv.id = "bot";
  botDiv.classList.add("message", "typing-animation");
  botDiv.style.animationDuration = "2s";
  botDiv.innerHTML = `<span id="bot-response">Typing...</span>`;
  mainDiv.appendChild(botDiv);

  var scroll = document.getElementById("chatBox");
  scroll.scrollTop = scroll.scrollHeight;
  inputField.disabled = true;
  quickMessageButtons.forEach(button => {
    button.disabled = true;
  });

  newUserDiv.classList.remove("typing-animation");
  setTimeout(() => {
    botDiv.classList.remove("typing-animation");
    botDiv.innerHTML = `<span id="bot-response">${product}</span>`;
    scroll.scrollTop = scroll.scrollHeight;
    inputField.disabled = false;
    quickMessageButtons.forEach(button => {
      button.disabled = false;
    });
    inputField.focus();
  }, 2000);
}
