import { GoogleGenAI, Type } from "https://esm.sh/@google/genai@1.39.0";
import * as docx from "https://esm.sh/docx@8.5.0";

// --- Constants ---
const BRAND_COLOR = "#ff4500";
const SYSTEM_INSTRUCTION = `You are an educational content generator for English exams.

A teacher will input a multiple-choice question about signs, notices, or messages.
The content inside square brackets [ ] is the primary subject.

Your task is a 3-step process:

Step 1: Classification
Identify the type based ONLY on the text inside [ ]:
- "sign": Short, rule-based, command or prohibition. Usually noun phrases or imperatives (e.g., NO PARKING, KEEP OFF).
- "notice": Informational message written in full sentences. Often includes time, reason, or future action (e.g., Closed for renovation. Will reopen next week).
- "message": A personal note, email, or handwritten text. Usually starts with a salutation (e.g., "Hi Mom", "Dear Tom") and ends with a name/signature. Informal tone.

Step 2: Image Description
Create a clear image description suitable for generating an illustration.
- If "sign": Describe a flat, symbolic design with strong visual icons. Minimal or no sentences. High contrast (e.g., Red circle for prohibition).
- If "notice": Describe a simple rectangular notice. You MUST explicitly state the exact text content to be written on the sign (e.g., "A notice with the text: 'Library Closed'"). Focus on the text content.
- If "message": Describe a handwritten note, sticky note, or piece of paper pinned to a surface. Mention that the font should look like legible HANDWRITING. Explicitly state the text content.

Step 3: Content Extraction
Prepare structured content for exporting to a DOCX file with the following layout:
Left side: the generated image
Right side: the question text and options A–D

Output your response in JSON format only, following this exact structure:
{
  "type": "sign" | "notice" | "message",
  "image_description": "...",
  "docx_content": {
    "question": "...",
    "options": {
      "A": "...",
      "B": "...",
      "C": "...",
      "D": "..."
    }
  }
}

Do NOT explain your reasoning.
Do NOT add extra text outside the JSON.`;

const REFERENCE_ANALYSIS_INSTRUCTION = `You are an assistant for an English teacher. 
The teacher will upload an image of a real-world Sign, Notice, or Message.

Your task:
1. Analyze the image to determine if it is a "Sign" (command/symbolic), a "Notice" (informational text), or a "Message" (handwritten note/email).
2. Generate 5 NEW, SIMILAR exam questions based on the style and context of the image.
3. Each suggestion must follow this exact format:
   "[Description of content] The question text...
    A. Option 1
    B. Option 2
    C. Option 3
    D. Option 4"

Return the result as a JSON Object containing an array of strings named "suggestions".`;

// --- DOM Elements ---
document.getElementById('currentYear').textContent = new Date().getFullYear();

// Tabs
const tabManual = document.getElementById('tabManual');
const tabUpload = document.getElementById('tabUpload');
const manualTabContent = document.getElementById('manualTabContent');
const uploadTabContent = document.getElementById('uploadTabContent');

// State
let selectedLevel = 'A2';
let activeTab = 'manual';
let referenceImageBase64 = null;
let referenceImageFile = null;
let suggestionsList = [];
let selectedIndices = new Set();
let batchResults = [];
let singleAnalysisResult = null;
let singleImageUrl = null;

// UI Elements
const levelA2Btn = document.getElementById('levelA2Btn');
const levelB1Btn = document.getElementById('levelB1Btn');
const questionInput = document.getElementById('questionInput');
const manualGenerateBtn = document.getElementById('manualGenerateBtn');
const apiKeyInput = document.getElementById('apiKeyInput');

// Error Box
const errorBox = document.getElementById('errorBox');
const errorMessageText = document.getElementById('errorMessageText');

// Upload Tab Elements
const fileInput = document.getElementById('fileInput');
const referenceImagePreview = document.getElementById('referenceImagePreview');
const refImgSrc = document.getElementById('refImgSrc');
const analyzingRefLoader = document.getElementById('analyzingRefLoader');
const suggestionsContainer = document.getElementById('suggestionsContainer');
const suggestionsListEl = document.getElementById('suggestionsList');
const selectAllBtn = document.getElementById('selectAllBtn');
const batchGenerateBtn = document.getElementById('batchGenerateBtn');
const batchProcessingLoader = document.getElementById('batchProcessingLoader');
const batchProgressText = document.getElementById('batchProgressText');
const batchResultsContainer = document.getElementById('batchResultsContainer');
const batchResultsCount = document.getElementById('batchResultsCount');
const batchResultsList = document.getElementById('batchResultsList');
const batchDownloadBtn = document.getElementById('batchDownloadBtn');
const startOverBtn = document.getElementById('startOverBtn');

// Single Result Elements
const singleResultContainer = document.getElementById('singleResultContainer');
const singleResultBadge = document.getElementById('singleResultBadge');
const singleDownloadBtn = document.getElementById('singleDownloadBtn');
const singleResultImgBox = document.getElementById('singleResultImgBox');
const singleResultImg = document.getElementById('singleResultImg');
const singleResultImgLoader = document.getElementById('singleResultImgLoader');
const singleResultDesc = document.getElementById('singleResultDesc');
const singleResultQuestion = document.getElementById('singleResultQuestion');
const singleResultOptions = document.getElementById('singleResultOptions');

// Helper
function showError(msg) {
  errorMessageText.textContent = msg;
  errorBox.classList.remove('hidden');
}

function hideError() {
  errorBox.classList.add('hidden');
}

function getApiKey() {
  const key = apiKeyInput.value.trim();
  if (!key) {
    showError("Please enter a Gemini API Key in the header.");
    throw new Error("Missing API Key");
  }
  return key;
}

// Helpers
const getBadgeStyle = (type) => {
  switch(type) {
    case 'sign': return 'bg-red-100 text-red-700';
    case 'notice': return 'bg-blue-100 text-blue-700';
    case 'message': return 'bg-green-100 text-green-700';
    default: return 'bg-gray-100 text-gray-700';
  }
};

const fileToGenerativePart = async (file) => {
  return new Promise((resolve) => {
    const reader = new FileReader();
    reader.onloadend = () => {
      resolve({
        inlineData: {
          data: reader.result.split(',')[1],
          mimeType: file.type,
        }
      });
    };
    reader.readAsDataURL(file);
  });
};

// Core Generation Logic
const generateQuestionItem = async (inputText, apiKey) => {
  const ai = new GoogleGenAI({ apiKey });

  const analysisResponse = await ai.models.generateContent({
    model: "gemini-3-flash-preview",
    contents: inputText,
    config: {
      systemInstruction: SYSTEM_INSTRUCTION,
      responseMimeType: "application/json",
      responseSchema: {
        type: Type.OBJECT,
        properties: {
          type: { type: Type.STRING, enum: ["sign", "notice", "message"] },
          image_description: { type: Type.STRING },
          docx_content: {
            type: Type.OBJECT,
            properties: {
              question: { type: Type.STRING },
              options: {
                type: Type.OBJECT,
                properties: {
                  A: { type: Type.STRING },
                  B: { type: Type.STRING },
                  C: { type: Type.STRING },
                  D: { type: Type.STRING },
                },
                required: ["A", "B", "C", "D"],
              },
            },
            required: ["question", "options"],
          },
        },
        required: ["type", "image_description", "docx_content"],
      },
    },
  });

  const jsonText = analysisResponse.text;
  if (!jsonText) throw new Error("No analysis response generated.");

  const result = JSON.parse(jsonText);

  const isNotice = result.type === "notice";
  const isMessage = result.type === "message";
  let imagePrompt = "";

  if (isNotice) {
    imagePrompt = `A high-contrast, black-and-white digital illustration of a TEXT NOTICE.
    Content: ${result.image_description}.
    
    STRICT VISUAL RULES FOR EXAM PRINTING:
    1. BACKGROUND: Pure white (#FFFFFF). No textures, no gradients, no shadows, no wall details.
    2. TEXT: Pitch black, bold, sans-serif font (Arial or Helvetica). MAXIMIZED SIZE and Center aligned.
    3. CONTAINER: A simple thin black rectangular border.
    4. STYLE: 2D Flat Vector. No photorealism.
    5. CLUTTER: ZERO clutter. No reflections, no glare.
    
    The goal is 100% readability for students reading this on a black-and-white exam paper.`;
  } else if (isMessage) {
    imagePrompt = `A high-contrast digital illustration of a HANDWRITTEN NOTE.
    Content: ${result.image_description}.
    
    STRICT VISUAL RULES FOR EXAM PRINTING:
    1. BACKGROUND: Pure white. REMOVE all background scenery (no corkboard, no fridge, no wall).
    2. CONTAINER: A simple white piece of paper with a subtle black outline.
    3. TEXT: Dark black "handwriting" style font. It must be perfectly LEGIBLE.
    4. STYLE: Minimalist line art / flat vector. 
    5. FOCUS: The text must be the main focus. No decorative elements.
    
    Ensure clear contrast so it prints well on paper.`;
  } else {
    imagePrompt = `A flat 2D vector icon of a SYMBOLIC SIGN: ${result.image_description}. 
    
    STRICT VISUAL RULES:
    1. BACKGROUND: Pure white.
    2. COLORS: Standard sign colors (Red, Blue, Black) only. High contrast.
    3. STYLE: Flat design, no shading, no gloss, no 3D effects. 
    4. FOCUS: The symbol/icon must be large and clear. No background noise.`;
  }

  const imageResponse = await ai.models.generateContent({
    model: "gemini-2.5-flash-image",
    contents: imagePrompt,
    config: {
      imageConfig: {
        aspectRatio: (isNotice || isMessage) ? "16:9" : "1:1",
      },
    },
  });

  let foundImageUrl = "";
  if (imageResponse.candidates && imageResponse.candidates[0].content.parts) {
    for (const part of imageResponse.candidates[0].content.parts) {
      if (part.inlineData) {
        const base64Data = part.inlineData.data;
        const mimeType = part.inlineData.mimeType || "image/png";
        foundImageUrl = `data:${mimeType};base64,${base64Data}`;
        break;
      }
    }
  }

  if (!foundImageUrl) {
    throw new Error("No image generated by the model.");
  }

  return { analysis: result, imageUrl: foundImageUrl };
};

// DOCX Creation
const createDocxBlob = async (items) => {
  const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, ImageRun } = docx;

  const sectionsChildren = [
    new Paragraph({
      children: [
        new TextRun({
          text: "English Exam - Sign Analysis",
          bold: true,
          size: 32, // 16pt
        }),
      ],
      spacing: { after: 400 },
    }),
  ];

  for (const item of items) {
    const response = await fetch(item.imageUrl);
    const blob = await response.blob();
    const arrayBuffer = await blob.arrayBuffer();
    const imageUint8Array = new Uint8Array(arrayBuffer);
    const mimeType = blob.type; 
    const docxImageType = mimeType.includes("jpeg") || mimeType.includes("jpg") ? "jpeg" : "png";
    const imageHeight = (item.analysis.type === "notice" || item.analysis.type === "message") ? 113 : 200;

    const itemTable = new Table({
      width: { size: 100, type: WidthType.PERCENTAGE },
      rows: [
        new TableRow({
          children: [
            new TableCell({
              width: { size: 45, type: WidthType.PERCENTAGE },
              children: [
                new Paragraph({
                  children: [
                    new ImageRun({
                      data: imageUint8Array,
                      transformation: { width: 200, height: imageHeight },
                      type: docxImageType,
                    }),
                  ],
                }),
              ],
            }),
            new TableCell({
              width: { size: 55, type: WidthType.PERCENTAGE },
              margins: { left: 200 },
              children: [
                new Paragraph({
                  children: [
                    new TextRun({
                      text: item.analysis.docx_content.question,
                      bold: true,
                      size: 24, // 12pt
                    }),
                  ],
                  spacing: { after: 200 },
                }),
                ...Object.entries(item.analysis.docx_content.options).map(
                  ([key, value]) =>
                    new Paragraph({
                      children: [
                        new TextRun({ text: `${key}. ${value}`, size: 24 }),
                      ],
                    })
                ),
              ],
            }),
          ],
        }),
      ],
    });

    sectionsChildren.push(itemTable);
    sectionsChildren.push(new Paragraph({ spacing: { before: 400 } }));
  }

  const doc = new Document({
    sections: [{ properties: {}, children: sectionsChildren }],
  });

  return await Packer.toBlob(doc);
};


// Actions
function switchTab(tab) {
  activeTab = tab;
  if (tab === 'manual') {
    tabManual.className = "flex-1 py-4 text-sm transition-all duration-200 border-b-2 border-[#ff4500] text-[#cf3700] bg-[#ffe0cc] font-bold shadow-[inset_0_-2px_4px_rgba(0,0,0,0.05)]";
    tabUpload.className = "flex-1 py-4 text-sm transition-all duration-200 border-b-2 border-transparent text-gray-500 hover:text-gray-700 hover:bg-gray-50 font-medium";
    manualTabContent.classList.remove('hidden');
    uploadTabContent.classList.add('hidden');
    // Also reset upload UI
  } else {
    tabUpload.className = "flex-1 py-4 text-sm transition-all duration-200 border-b-2 border-[#ff4500] text-[#cf3700] bg-[#ffe0cc] font-bold shadow-[inset_0_-2px_4px_rgba(0,0,0,0.05)]";
    tabManual.className = "flex-1 py-4 text-sm transition-all duration-200 border-b-2 border-transparent text-gray-500 hover:text-gray-700 hover:bg-gray-50 font-medium";
    uploadTabContent.classList.remove('hidden');
    manualTabContent.classList.add('hidden');
    
    // Hide single result
    singleResultContainer.classList.add('hidden');
  }
}

tabManual.addEventListener('click', () => switchTab('manual'));
tabUpload.addEventListener('click', () => switchTab('upload'));

function setLevel(lvl) {
  selectedLevel = lvl;
  if (lvl === 'A2') {
    levelA2Btn.className = "px-5 py-2.5 rounded-md text-sm font-semibold transition-all duration-200 bg-white text-[#ff4500] shadow-sm ring-1 ring-gray-200";
    levelB1Btn.className = "px-5 py-2.5 rounded-md text-sm font-semibold transition-all duration-200 text-gray-500 hover:text-gray-700";
  } else {
    levelB1Btn.className = "px-5 py-2.5 rounded-md text-sm font-semibold transition-all duration-200 bg-white text-[#ff4500] shadow-sm ring-1 ring-gray-200";
    levelA2Btn.className = "px-5 py-2.5 rounded-md text-sm font-semibold transition-all duration-200 text-gray-500 hover:text-gray-700";
  }
}

levelA2Btn.addEventListener('click', () => setLevel('A2'));
levelB1Btn.addEventListener('click', () => setLevel('B1'));

// --- Manual Generate ---
manualGenerateBtn.addEventListener('click', async () => {
  const text = questionInput.value.trim();
  if (!text) return;
  hideError();
  
  try {
    const apiKey = getApiKey();
    manualGenerateBtn.textContent = "Classifying & Analyzing...";
    manualGenerateBtn.disabled = true;
    manualGenerateBtn.classList.replace('bg-[#ff4500]', 'bg-gray-400');
    manualGenerateBtn.classList.add('cursor-not-allowed');

    singleResultContainer.classList.add('hidden');
    singleResultImg.classList.add('hidden');
    singleResultImgLoader.classList.remove('hidden');
    singleResultOptions.innerHTML = '';

    const { analysis, imageUrl } = await generateQuestionItem(text, apiKey);
    
    singleAnalysisResult = analysis;
    singleImageUrl = imageUrl;

    // Populate UI
    singleResultBadge.className = `px-3 py-1 text-xs font-bold uppercase tracking-wider rounded-full ${getBadgeStyle(analysis.type)}`;
    singleResultBadge.textContent = `Detected: ${analysis.type}`;
    
    if (analysis.type === "notice" || analysis.type === "message") {
      singleResultImgBox.classList.replace('aspect-square', 'aspect-video');
    } else {
      singleResultImgBox.classList.replace('aspect-video', 'aspect-square');
    }

    singleResultImg.src = imageUrl;
    singleResultImg.classList.remove('hidden');
    singleResultImgLoader.classList.add('hidden');
    
    singleResultDesc.textContent = `"${analysis.image_description}"`;
    singleResultQuestion.textContent = analysis.docx_content.question;

    Object.entries(analysis.docx_content.options).forEach(([key, val]) => {
      const div = document.createElement('div');
      div.className = "flex items-start group";
      div.innerHTML = `
        <span class="flex-shrink-0 w-8 h-8 rounded-full bg-[#ff4500] text-white flex items-center justify-center font-bold text-sm mr-4 shadow-sm group-hover:bg-[#e03e00] transition-colors">${key}</span>
        <span class="text-lg text-gray-700 pt-0.5">${val}</span>
      `;
      singleResultOptions.appendChild(div);
    });

    singleResultContainer.classList.remove('hidden');
  } catch (err) {
    console.error(err);
    showError(err.message || "An unexpected error occurred.");
  } finally {
    manualGenerateBtn.textContent = "Generate Content";
    manualGenerateBtn.disabled = false;
    manualGenerateBtn.classList.replace('bg-gray-400', 'bg-[#ff4500]');
    manualGenerateBtn.classList.remove('cursor-not-allowed');
  }
});

// Download Single
singleDownloadBtn.addEventListener('click', async () => {
  if (!singleAnalysisResult || !singleImageUrl) return;
  try {
    const blob = await createDocxBlob([{ analysis: singleAnalysisResult, imageUrl: singleImageUrl }]);
    const url = window.URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = "exam-question.docx";
    a.click();
    window.URL.revokeObjectURL(url);
  } catch (e) {
    showError("Failed to create DOCX file.");
  }
});

// --- Upload Workflow ---
fileInput.addEventListener('change', async (e) => {
  const file = e.target.files[0];
  if (!file) return;

  hideError();
  referenceImageFile = file;
  
  const reader = new FileReader();
  reader.onload = (ev) => {
    referenceImageBase64 = ev.target.result;
    refImgSrc.src = referenceImageBase64;
    referenceImagePreview.classList.remove('hidden');
  };
  reader.readAsDataURL(file);

  suggestionsContainer.classList.add('hidden');
  batchResultsContainer.classList.add('hidden');
  suggestionsList = [];
  selectedIndices.clear();
  batchResults = [];

  try {
    const apiKey = getApiKey();
    analyzingRefLoader.classList.remove('hidden');

    const ai = new GoogleGenAI({ apiKey });
    const imagePart = await fileToGenerativePart(file);
    
    const levelCriteria = selectedLevel === "A2" 
      ? "A2: simple vocabulary, direct meanings, minimal inference."
      : "B1: slightly longer sentences, simple inference, clearer distractors.";

    const promptText = `Analyze this image and generate 5 similar exam questions.
    
    Strictly follow these requirements:
    1. Target Level: ${selectedLevel}
    2. Level Criteria: ${levelCriteria}
    3. Match the identified type (sign, notice, or message).
    4. Match the exam style of the input.
    5. Do NOT copy wording or meaning from the example. Create NEW content.`;

    const response = await ai.models.generateContent({
      model: "gemini-3-flash-preview",
      contents: { parts: [ imagePart, { text: promptText } ] },
      config: {
        systemInstruction: REFERENCE_ANALYSIS_INSTRUCTION,
        responseMimeType: "application/json",
        responseSchema: {
          type: Type.OBJECT,
          properties: {
            suggestions: {
              type: Type.ARRAY,
              items: { type: Type.STRING }
            }
          }
        }
      }
    });

    const result = JSON.parse(response.text || "{}");
    if (result.suggestions && Array.isArray(result.suggestions)) {
      suggestionsList = result.suggestions;
      renderSuggestions();
    } else {
      throw new Error("Failed to generate suggestions.");
    }
  } catch (err) {
    console.error(err);
    showError("Failed to analyze image. " + (err.message || ""));
  } finally {
    analyzingRefLoader.classList.add('hidden');
  }
});

function renderSuggestions() {
  suggestionsContainer.classList.remove('hidden');
  suggestionsListEl.innerHTML = '';
  
  suggestionsList.forEach((sug, i) => {
    const isSelected = selectedIndices.has(i);
    const div = document.createElement('div');
    
    div.className = `p-4 rounded-lg border transition-all cursor-pointer flex items-start gap-4 ${isSelected ? "border-[#ff4500] bg-white shadow-sm ring-1 ring-[#ff4500]" : "border-gray-200 bg-white hover:bg-gray-50"}`;
    div.innerHTML = `
      <div class="mt-1 flex-shrink-0 w-5 h-5 rounded border flex items-center justify-center ${isSelected ? "bg-[#ff4500] border-[#ff4500]" : "border-gray-300 bg-white"}">
        ${isSelected ? '<svg class="w-3 h-3 text-white" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="3" d="M5 13l4 4L19 7"></path></svg>' : ''}
      </div>
      <div class="flex-1">
         <pre class="text-sm text-gray-700 whitespace-pre-wrap font-sans leading-relaxed">${sug}</pre>
         <div class="mt-2 text-right">
           <button class="use-single-btn text-xs font-semibold text-[#ff4500] hover:underline" data-idx="${i}">Edit Individually</button>
         </div>
      </div>
    `;

    // Click on row to toggle selection
    div.addEventListener('click', () => {
      if (selectedIndices.has(i)) selectedIndices.delete(i);
      else selectedIndices.add(i);
      renderSuggestions();
    });

    // Click on inner button
    div.querySelector('.use-single-btn').addEventListener('click', (e) => {
      e.stopPropagation();
      questionInput.value = sug;
      switchTab('manual');
      
      // reset
      suggestionsList = [];
      referenceImageBase64 = null;
      selectedIndices.clear();
      fileInput.value = "";
      suggestionsContainer.classList.add('hidden');
      referenceImagePreview.classList.add('hidden');
    });

    suggestionsListEl.appendChild(div);
  });

  updateBatchBtn();
}

function updateBatchBtn() {
  if (selectedIndices.size > 0) {
    batchGenerateBtn.disabled = false;
    batchGenerateBtn.classList.remove('bg-gray-400', 'cursor-not-allowed');
    batchGenerateBtn.classList.add('bg-[#ff4500]', 'hover:bg-[#e03e00]', 'hover:shadow-lg', 'active:scale-95');
    batchGenerateBtn.textContent = `Generate ${selectedIndices.size} Selected Questions`;
    selectAllBtn.textContent = selectedIndices.size === suggestionsList.length ? "Deselect All" : "Select All";
  } else {
    batchGenerateBtn.disabled = true;
    batchGenerateBtn.classList.add('bg-gray-400', 'cursor-not-allowed');
    batchGenerateBtn.classList.remove('bg-[#ff4500]', 'hover:bg-[#e03e00]', 'hover:shadow-lg', 'active:scale-95');
    batchGenerateBtn.textContent = `Generate Selected Questions`;
    selectAllBtn.textContent = "Select All";
  }
}

selectAllBtn.addEventListener('click', () => {
  if (selectedIndices.size === suggestionsList.length) {
    selectedIndices.clear();
  } else {
    suggestionsList.forEach((_, i) => selectedIndices.add(i));
  }
  renderSuggestions();
});


// Batch Generation
batchGenerateBtn.addEventListener('click', async () => {
  if (selectedIndices.size === 0) return;
  hideError();

  suggestionsContainer.classList.add('hidden');
  batchProcessingLoader.classList.remove('hidden');
  batchResults = [];

  try {
    const apiKey = getApiKey();
    const indices = Array.from(selectedIndices).sort((a,b) => a-b);
    
    for (let i = 0; i < indices.length; i++) {
       const idx = indices[i];
       const text = suggestionsList[idx];
       batchProgressText.textContent = `Generating question ${i + 1} of ${indices.length}...`;

       const { analysis, imageUrl } = await generateQuestionItem(text, apiKey);
       batchResults.push({ id: idx, analysis, imageUrl });
    }

    renderBatchResults();
  } catch(err) {
    console.error(err);
    showError("Batch Error: " + (err.message || ""));
    suggestionsContainer.classList.remove('hidden');
  } finally {
    batchProcessingLoader.classList.add('hidden');
  }
});

function renderBatchResults() {
  batchResultsContainer.classList.remove('hidden');
  batchResultsCount.textContent = `Batch Results (${batchResults.length})`;
  batchResultsList.innerHTML = '';
  
  batchResults.forEach((result, i) => {
    const div = document.createElement('div');
    div.className = "bg-white p-6 border border-gray-200 rounded-xl shadow-md";
    div.innerHTML = `
      <div class="flex flex-col md:flex-row gap-8">
        <div class="w-full md:w-1/3 flex-shrink-0">
          <div class="bg-gray-100 rounded-lg border border-gray-200 aspect-video flex items-center justify-center p-4">
             <img src="${result.imageUrl}" class="w-full h-full object-contain drop-shadow-sm" />
          </div>
        </div>
        <div class="w-full md:w-2/3">
          <div class="flex items-center gap-3 mb-3">
            <span class="bg-gray-800 text-white text-xs font-bold px-2 py-1 rounded">Q${i+1}</span>
            <span class="text-xs font-bold uppercase tracking-wider px-2 py-1 rounded ${getBadgeStyle(result.analysis.type)}">${result.analysis.type}</span>
          </div>
          <p class="font-bold text-lg text-gray-900 mb-4">${result.analysis.docx_content.question}</p>
          <div class="grid grid-cols-1 sm:grid-cols-2 gap-3">
             <div class="bg-gray-50 p-3 rounded border border-gray-100 text-sm"><span class="font-bold text-[#ff4500]">A.</span> ${result.analysis.docx_content.options.A}</div>
             <div class="bg-gray-50 p-3 rounded border border-gray-100 text-sm"><span class="font-bold text-[#ff4500]">B.</span> ${result.analysis.docx_content.options.B}</div>
             <div class="bg-gray-50 p-3 rounded border border-gray-100 text-sm"><span class="font-bold text-[#ff4500]">C.</span> ${result.analysis.docx_content.options.C}</div>
             <div class="bg-gray-50 p-3 rounded border border-gray-100 text-sm"><span class="font-bold text-[#ff4500]">D.</span> ${result.analysis.docx_content.options.D}</div>
          </div>
        </div>
      </div>
    `;
    batchResultsList.appendChild(div);
  });
}

batchDownloadBtn.addEventListener('click', async () => {
  if (batchResults.length === 0) return;
  try {
    const docData = batchResults.map(r => ({ analysis: r.analysis, imageUrl: r.imageUrl }));
    const blob = await createDocxBlob(docData);
    const url = window.URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = "exam-questions-batch.docx";
    a.click();
    window.URL.revokeObjectURL(url);
  } catch (e) {
    showError("Failed to create Batch DOCX file.");
  }
});

startOverBtn.addEventListener('click', () => {
  batchResults = [];
  selectedIndices.clear();
  suggestionsList = [];
  referenceImageBase64 = null;
  fileInput.value = "";
  
  batchResultsContainer.classList.add('hidden');
  referenceImagePreview.classList.add('hidden');
  uploadTabContent.classList.remove('hidden'); // ensure upload tab is showing the dropzone
});
