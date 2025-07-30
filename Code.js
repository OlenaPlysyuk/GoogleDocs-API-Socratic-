
// Constants

// Enter your key there if in Apps Script

require('dotenv').config();//="YOUR KEY"; if in APPS SCRIPT


const API_KEY = process.env.API_KEY;//remove this if in Apps Script
const MODEL_TYPE = "gpt-4o-mini";

// One-time system instruction for the assistant
const SYSTEM_PROMPT =
 "You are an expert limerick writer helping a student learn to write limericks using the Socratic method. Never directly provide the answer or write the student's limerick for them. Your responses must consist of carefully guided questions or hints …";


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
 return reply;
}

// ------------------------------  SIDEBAR  ------------------------------

function showSidebar() {
 const html = HtmlService.createHtmlOutputFromFile('Sidebar').setTitle('LimerickTool');
 DocumentApp.getUi().showSidebar(html);
}
function processPromptInsert(prompt) { if (prompt) callOpenAI_(prompt, true); }
function processPromptShow(prompt)   { return prompt ? callOpenAI_(prompt, false) : ''; }

// ------------------------------  SYLLABLE UTILITIES  ------------------------------
/*
function countSyllablesWord_(w) {
 w = w.toLowerCase().replace(/[^a-z]/g, '');
 if (!w)            return 0;
 if (w.length <= 3) return 1;
 w = w.replace(/(?:[^laeiouy]es|ed|[^laeiouy]e)$/ , '');
 w = w.replace(/^y/, '');
 const m = w.match(/[aeiouy]{1,2}/g);
 return m ? m.length : 1;
}
function countSyllablesString_(s) {
 return s.split(/\s+/).reduce((acc, w) => acc + countSyllablesWord_(w), 0);
}
function superscriptDigits_(n) {
 const map = { '0':'⁰','1':'¹','2':'²','3':'³','4':'⁴','5':'⁵','6':'⁶','7':'⁷','8':'⁸','9':'⁹' };
 return String(n).split('').map(d => map[d] || d).join('');
}
*/
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

// ------------------------------  SYLLABLE ANNOTATION  ------------------------------
/*
function annotateSyllables() {
 const ui  = DocumentApp.getUi();
 const sel = DocumentApp.getActiveDocument().getSelection();
 if (!sel) return ui.alert('Please select word(s) or line(s) first.');

 const plain = _extractSelectionPlainText_();
 if (!plain || !plain.trim()) return ui.alert('Selection contains no text.');

 const isSingleWord = !/[\s\n]/.test(plain.trim());

 const firstRange = sel.getRangeElements()[0];
 const firstText  = firstRange.getElement().asText();
 let   sOff       = firstRange.getStartOffset();
 if (sOff === -1) sOff = 0;

 if (isSingleWord) {
   const syl  = countSyllablesWord_(plain);
   const mark = superscriptDigits_(syl) + ' ';
   firstText.insertText(sOff, mark);
   firstText.setAttributes(sOff, sOff + mark.length - 1, {
     [DocumentApp.Attribute.FONT_SIZE]        : 9,
     [DocumentApp.Attribute.FOREGROUND_COLOR] : '#666666'
   });
 } else {
   const syl   = countSyllablesString_(plain);
   const label = ' [' + syl + ' syllables]';
   const insertPos = firstText.getText().length;
   firstText.appendText(label);
   firstText.setAttributes(insertPos, insertPos + label.length - 1, {
     [DocumentApp.Attribute.ITALIC]           : true,
     [DocumentApp.Attribute.FONT_SIZE]        : 9,
     [DocumentApp.Attribute.FOREGROUND_COLOR] : '#666666'
   });
 }
}
*/
// ------------------------------  CLEAR SYLLABLE ANNOTATIONS  ------------------------------
/*
function clearSyllableAnnotations() {
 const body = DocumentApp.getActiveDocument().getBody();
 const txt  = body.editAsText();
 if (!txt) return;

 //1. Remove superscript digits ⁰–⁹ followed by optional whitespace (incl NBSP) 
 txt.replaceText('[⁰¹²³⁴⁵⁶⁷⁸⁹]+(?:\s|\u00A0)*', '');

 // 2. Remove labels like [7 syllables] or [7 syllable] 
 txt.replaceText('\\s*\\[\\s*\\d+\\s*syllables?\\s*\\]', '');
}

*/
// ------------------------------  RHYME UTILITIES  ------------------------------
/*
function safeParseJson_(txt) {
 try {
   return JSON.parse(txt);
 } catch (e) {
   try {
     const core = txt.slice(txt.indexOf('{'), txt.lastIndexOf('}') + 1).replace(/,\s*([\]}])/g, '$1');
     return JSON.parse(core);
   } catch (err) {
     return { rhymes: [], non_rhymes: [] };
   }
 }
}

function _clusterRhymesByEnding_(words) {
 const groups = {};
 words.forEach(w => {
   const key = w.toLowerCase().replace(/[^a-z]/g, '').slice(-3);
   if (!key) return;
   (groups[key] = groups[key] || []).push(w);
 });
 return Object.values(groups);
}

function getRhymeSets_(poem) {
 const prompt = 'Return ONLY valid JSON {"rhymes": [...], "non_rhymes": [...]} for poem:\n"""' + poem + '"""';
 const payload = {
   model: MODEL_TYPE,
   messages: [{ role: 'user', content: prompt }],
   temperature: 0.2,
   max_tokens: 120,
   response_format: { type: 'json_object' }
 };

 const resp = UrlFetchApp.fetch('https://api.openai.com/v1/chat/completions', {
   method:      'post',
   contentType: 'application/json',
   headers:     { Authorization: 'Bearer ' + API_KEY },
   payload:     JSON.stringify(payload)
 });

 const raw  = JSON.parse(resp.getContentText()).choices[0].message.content.trim();
 return safeParseJson_(raw);
}
*/
// ------------------------------  HIGHLIGHT HELPERS  ------------------------------
/*
function highlightWords_(words, color) {
 if (!words || !words.length) return;
 const body = DocumentApp.getActiveDocument().getBody();
 const esc  = s => s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
 words.forEach(w => {
   const pattern = '(?i)\\b' + esc(w) + '\\b[.,;:!?]?';
   let m = body.findText(pattern);
   while (m) {
     const el = m.getElement();
     let s = m.getStartOffset();
     let e = m.getEndOffsetInclusive();
     const txt = el.getText();
     if (/[.,;:!?]/.test(txt.charAt(e + 1))) e += 1; // capture punctuation
     el.setBackgroundColor(s, e, color);
     m = body.findText(pattern, m);
   }
 });
}

// Helper to generate visually distinct colors on the fly (HSL → HEX)
function hueToHex_(h) {
 const a = (h + 360) % 360 / 360;
 const f = n => {
   const k = (n + a * 12) % 12;
   const c = 0.65 - Math.max(Math.min(k - 3, 9 - k, 1), -1) * 0.35;
   return Math.round(c * 255).toString(16).padStart(2, '0');
 };
 return '#' + f(0) + f(8) + f(4);
}

function highlightRhymes() {
 const doc   = DocumentApp.getActiveDocument();
 const plain = _extractSelectionPlainText_() || doc.getBody().getText();
 const sets  = getRhymeSets_(plain);

 // --- 1. Normalize rhyme groups -------------------------------------------
 let rhymeGroups = sets.rhymes || [];
 if (rhymeGroups.length && typeof rhymeGroups[0] === 'string') {
   rhymeGroups = _clusterRhymesByEnding_(rhymeGroups);
 }

 // --- 2. Highlight each group with its own color --------------------------
 let idx = 0;
 rhymeGroups.forEach(grp => {
   const color = hueToHex_(idx * 137.508); // Golden‑angle increment for diversity
   idx += 1;
   highlightWords_(grp, color);
 });

 // --- 3. Highlight lonely end‑words (no rhyme found) -----------------------
 const lastWords = [];
 plain.split(/\n+/).forEach(line => {
   const m = line.trim().match(/[A-Za-z']+$/);
   if (m) lastWords.push(m[0]);
 });
 const inRhyme = new Set(rhymeGroups.flat().map(w => w.toLowerCase()));
 const lonely  = lastWords.filter(w => !inRhyme.has(w.toLowerCase()));
 highlightWords_(lonely, '#F8BBD0'); // soft pink for lonely endings

 // --- 4. Non‑rhymes list from model ---------------------------------------
 highlightWords_(sets.non_rhymes || [], '#E0E0E0');
}

// ------------------------------  CLEAR HIGHLIGHTS  ------------------------------

//function clearHighlights() {
// const body = DocumentApp.getActiveDocument().getBody();
 //body.getParagraphs().forEach(par => {
 //  for (let i = 0; i < par.getNumChildren(); i++) {
    // const el = par.getChild(i);
    // if (el.getType() === DocumentApp.ElementType.TEXT) {
    //   el.asText().setBackgroundColor(null);
    // }
   //}
 //});
//}
*/
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
     .addItem('Generate prompt (selection)', 'generatePromptFromSelection')
     .addSeparator()
     .addToUi();
     //.addItem('Highlight rhymes',            'highlightRhymes')
     //addItem('Clear rhyme highlights',      'clearHighlights')
     //.addSeparator()
     //.addItem('Annotate syllables',          'annotateSyllables')
     //.addItem('Clear syllable marks',        'clearSyllableAnnotations')
     // .addSeparator()
     //.addItem('Clear chat history',          'clearChat')
    
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

   // Витягаємо слова, відфільтровуємо фрази й дублікати
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
  // const sheet = SpreadsheetApp.openById('ENETER YOUR SHEET ID THERE').getSheetByName('ChatLog'); UNCOMMENT THIS
  sheet.appendRow([new Date(), actionType, JSON.stringify(payload)]);
}

// ------------------------------  END  ------------------------------


