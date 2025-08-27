
// Constants

const API_KEY = PropertiesService.getScriptProperties().getProperty('OPENAI_KEY');
const MODEL_TYPE = "gpt-4o-mini";
const SYSTEM_PROMPT =
 "You are an expert limerick coach using a supportive Socratic method. Never provide full lines or the completed limerick. Guide the student with questions, hints, and partial ideas—such as rhymes, thematic suggestions, or imagery prompts—while keeping them responsible for crafting the lines. You may quote their own text to point out rhythm, meter, or rhyme issues, and suggest words or directions, but always encourage them to experiment and decide.";



// ------------------------------  CHAT HISTORY HELPERS  ------------------------------

function loadHistory_() {
 const raw = PropertiesService.getDocumentProperties().getProperty('CHAT_HISTORY');
 return raw ? JSON.parse(raw) : [];
}
function saveHistory_(h) {
 PropertiesService.getDocumentProperties().setProperty('CHAT_HISTORY', JSON.stringify(h));
}

// ------------------------------  CORE OPENAI CALL  ------------------------------

function callOpenAI_(userText, insertIntoDoc) {
 let hist = loadHistory_();
 if (!hist.length) hist.push({ role: 'system', content: SYSTEM_PROMPT });
 hist.push({ role: 'user', content: userText });
 if (hist.length > 50) hist = [hist[0], ...hist.slice(-50)];

 const payload = {
   model: MODEL_TYPE,
   messages: hist,
   temperature: 0.5,
   max_tokens: 256
 };

 const resp = UrlFetchApp.fetch('https://api.openai.com/v1/chat/completions', {
   method:      'post',
   contentType: 'application/json',
   headers:     { Authorization: 'Bearer ' + API_KEY },
   payload:     JSON.stringify(payload)
 });

 const reply = JSON.parse(resp.getContentText()).choices[0].message.content.trim();

  hist.push({ role: 'assistant', content: reply });
 saveHistory_(hist);

 logToSheet('user_prompt', userText);
 logToSheet('assistant_reply', reply);

 if (insertIntoDoc) {
   DocumentApp.getActiveDocument().getBody().appendParagraph(reply).setItalic(true);
 }
 //logToSheet('user_prompt', userText);
 //logToSheet('assistant_reply', reply);
 return reply;


}

// ------------------------------  SIDEBAR  ------------------------------

function showSidebar() {
 const html = HtmlService.createHtmlOutputFromFile('Sidebar').setTitle('LimerickTool');
 DocumentApp.getUi().showSidebar(html);
}
function processPromptInsert(prompt) { if (prompt) callOpenAI_(prompt, true); }
function processPromptShow(prompt)   { return prompt ? callOpenAI_(prompt, false) : ''; }


// ------------------------------  SELECTION HELPER  ------------------------------

function _extractSelectionPlainText_() {
 const sel = DocumentApp.getActiveDocument().getSelection();
 if (!sel) return null;
 const txtArr = [];
 sel.getRangeElements().forEach(re => {
   const el = re.getElement();
   if (!el.editAsText) return;
   const t    = el.asText();
   const full = t.getText();
   let sOff = re.getStartOffset();
   let eOff = re.getEndOffsetInclusive();
   if (sOff === -1 || eOff === -1) { sOff = 0; eOff = full.length - 1; }
   txtArr.push(full.substring(sOff, eOff + 1));
 });
 return txtArr.join('\n');
}


// ------------------------------  PROMPT HELPERS  ------------------------------

function generatePromptFromSelection() {
 const ui    = DocumentApp.getUi();
 const plain = _extractSelectionPlainText_();
 if (!plain) return ui.alert('Please select some text first.');
 callOpenAI_(plain.trim(), true);
}

function clearChat() {
 PropertiesService.getDocumentProperties().deleteProperty('CHAT_HISTORY');
 DocumentApp.getUi().alert('Chat history cleared.');
}

function getChatHistory() {
  const raw = PropertiesService.getDocumentProperties().getProperty('CHAT_HISTORY');
  const hist = raw ? JSON.parse(raw) : [];
  return hist.filter(e => e.role === 'user' || e.role === 'assistant');
}

// ------------------------------  onOpen MENU  ------------------------------

function onOpen() {
 DocumentApp.getUi()
   .createMenu('Limericks Tool')
     .addItem('Open chat sidebar',           'showSidebar')
     .addSeparator()

     .addToUi();
     
    
}

// ------------------------------  SINGLE-WORD RHYME LOOKUP  ------------------------------

function findRhymes(word) {
 if (!word) return [];

 try {
  const url   = 'https://api.datamuse.com/words?rel_rhy=' +
                encodeURIComponent(word.toLowerCase()) +
               '&max=50';
   const resp  = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
  const items = JSON.parse(resp.getContentText()) || [];

   // Filter duplicates and getting words
  const seen = new Set();
   const rhymes = [];
   items.forEach(it => {
     const w = (it.word || '').trim();
     if (w && !w.includes(' ') && !seen.has(w)) {
       seen.add(w);
       rhymes.push(w);
     }
   });

   logToSheet('rhyme_lookup', { word: word, result: rhymes });
 return rhymes;
 } catch (e) {
   return [];
}
 }

// ------------------------------  SAVING LOGS  ------------------------------

function logToSheet(actionType, payload) {
// const sheet = SpreadsheetApp.openById('ENTER YOUR ID HERE').getSheetByName('SocraticMethodLogs');
 sheet.appendRow([new Date(), actionType, JSON.stringify(payload)]);
}

// ------------------------------  END  ------------------------------


