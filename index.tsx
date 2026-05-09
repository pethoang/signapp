import React, { useState, useRef } from "react";
import { createRoot } from "react-dom/client";
import { GoogleGenAI, Type } from "@google/genai";
import * as docx from "https://esm.sh/docx@8.5.0";

// --- Types ---

interface DocxContent {
  question: string;
  options: {
    A: string;
    B: string;
    C: string;
    D: string;
  };
}

interface AnalysisResult {
  type: "sign" | "notice" | "message";
  image_description: string;
  docx_content: DocxContent;
}

interface BatchItem {
  id: number;
  analysis: AnalysisResult;
  imageUrl: string;
}

// --- Constants ---

const BRAND_COLOR = "#ff4500"; // Primary Orange Red

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

// --- Helper: Core Generation Logic ---
const generateQuestionItem = async (inputText: string, apiKey: string): Promise<{ analysis: AnalysisResult, imageUrl: string }> => {
  const ai = new GoogleGenAI({ apiKey });

  // 1. Analyze Text & Classify
  const analysisModel = "gemini-3-flash-preview";
  const analysisResponse = await ai.models.generateContent({
    model: analysisModel,
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

  const result: AnalysisResult = JSON.parse(jsonText);

  // 2. Generate Image based on Type
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

  const imageModel = "gemini-2.5-flash-image";
  const imageResponse = await ai.models.generateContent({
    model: imageModel,
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

// --- App Component ---

const App = () => {
  // Tabs: 'manual' is the existing flow, 'upload' is the new feature
  const [activeTab, setActiveTab] = useState<"manual" | "upload">("manual");

  // Existing State (Manual Mode)
  const [inputText, setInputText] = useState<string>(
    "[The garden is closed for planting. Will be open to visitors next week] The notice says...\nA. You can visit anytime.\nB. Visitors are allowed now.\nC. Closed for planting.\nD. Open after planting."
  );
  const [loadingStep, setLoadingStep] = useState<
    "idle" | "analyzing" | "drawing" | "done"
  >("idle");
  const [analysisResult, setAnalysisResult] = useState<AnalysisResult | null>(
    null
  );
  const [imageUrl, setImageUrl] = useState<string | null>(null);
  const [error, setError] = useState<string | null>(null);

  // New State for Upload Feature & Batch
  const [referenceImage, setReferenceImage] = useState<string | null>(null);
  const [isAnalyzingRef, setIsAnalyzingRef] = useState(false);
  const [suggestions, setSuggestions] = useState<string[]>([]);
  const [selectedIndices, setSelectedIndices] = useState<Set<number>>(new Set());
  const [batchResults, setBatchResults] = useState<BatchItem[]>([]);
  const [isBatchProcessing, setIsBatchProcessing] = useState(false);
  const [batchProgress, setBatchProgress] = useState("");
  const [selectedLevel, setSelectedLevel] = useState<"A2" | "B1">("A2");
  
  const fileInputRef = useRef<HTMLInputElement>(null);

  // --- Helper: Convert Blob to Base64 ---
  const fileToGenerativePart = async (file: File) => {
    const base64EncodedDataPromise = new Promise<string>((resolve) => {
      const reader = new FileReader();
      reader.onloadend = () => resolve((reader.result as string).split(",")[1]);
      reader.readAsDataURL(file);
    });
    return {
      inlineData: {
        data: await base64EncodedDataPromise,
        mimeType: file.type,
      },
    };
  };

  // --- Feature: Analyze Reference Image ---
  const handleReferenceUpload = async (event: React.ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (e) => setReferenceImage(e.target?.result as string);
    reader.readAsDataURL(file);

    setIsAnalyzingRef(true);
    setSuggestions([]);
    setSelectedIndices(new Set());
    setBatchResults([]);
    setError(null);

    try {
      const apiKey = process.env.API_KEY;
      if (!apiKey) throw new Error("API Key not found.");

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
        contents: {
          parts: [
            imagePart,
            { text: promptText }
          ]
        },
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
        setSuggestions(result.suggestions);
      } else {
        throw new Error("Failed to generate suggestions.");
      }

    } catch (err: any) {
      console.error(err);
      setError("Failed to analyze image. Please try again. " + (err.message || ""));
    } finally {
      setIsAnalyzingRef(false);
    }
  };

  const toggleSelection = (index: number) => {
    const newSet = new Set(selectedIndices);
    if (newSet.has(index)) {
      newSet.delete(index);
    } else {
      newSet.add(index);
    }
    setSelectedIndices(newSet);
  };

  const handleSelectAll = () => {
    if (selectedIndices.size === suggestions.length) {
      setSelectedIndices(new Set());
    } else {
      const newSet = new Set<number>();
      suggestions.forEach((_, i) => newSet.add(i));
      setSelectedIndices(newSet);
    }
  };

  const handleUseSingleSuggestion = (suggestion: string) => {
    setInputText(suggestion);
    setActiveTab("manual");
    setSuggestions([]);
    setReferenceImage(null);
    setSelectedIndices(new Set());
    if (fileInputRef.current) fileInputRef.current.value = "";
  };

  // --- Feature: Batch Generation ---
  const handleBatchGenerate = async () => {
    if (selectedIndices.size === 0) return;

    setIsBatchProcessing(true);
    setBatchResults([]);
    setError(null);

    const apiKey = process.env.API_KEY;
    if (!apiKey) {
      setError("API Key not found.");
      setIsBatchProcessing(false);
      return;
    }

    const indicesToProcess = (Array.from(selectedIndices) as number[]).sort((a, b) => a - b);
    const results: BatchItem[] = [];

    try {
      for (let i = 0; i < indicesToProcess.length; i++) {
        const index = indicesToProcess[i];
        const text = suggestions[index];
        setBatchProgress(`Generating question ${i + 1} of ${indicesToProcess.length}...`);

        const { analysis, imageUrl } = await generateQuestionItem(text, apiKey);
        
        results.push({
          id: index,
          analysis,
          imageUrl
        });
        
        setBatchResults([...results]);
      }
    } catch (err: any) {
      console.error(err);
      setError("Error during batch generation: " + (err.message || "Unknown error"));
    } finally {
      setIsBatchProcessing(false);
      setBatchProgress("");
    }
  };

  // --- Existing Logic: Generate Single Content ---
  const handleGenerate = async () => {
    if (!inputText.trim()) return;

    setLoadingStep("analyzing");
    setError(null);
    setAnalysisResult(null);
    setImageUrl(null);

    try {
      const apiKey = process.env.API_KEY;
      if (!apiKey) throw new Error("API Key not found in environment.");

      const { analysis, imageUrl: url } = await generateQuestionItem(inputText, apiKey);
      
      setAnalysisResult(analysis);
      setLoadingStep("drawing");
      setImageUrl(url);
      setLoadingStep("done");

    } catch (err: any) {
      console.error(err);
      setError(err.message || "An unexpected error occurred.");
      setLoadingStep("idle");
    }
  };

  // --- DOCX Generation Logic ---
  const createDocxBlob = async (items: Array<{analysis: AnalysisResult, imageUrl: string}>) => {
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
        width: {
          size: 100,
          type: WidthType.PERCENTAGE,
        },
        rows: [
          new TableRow({
            children: [
              new TableCell({
                width: {
                  size: 45,
                  type: WidthType.PERCENTAGE,
                },
                children: [
                  new Paragraph({
                    children: [
                      new ImageRun({
                        data: imageUint8Array,
                        transformation: {
                          width: 200,
                          height: imageHeight, 
                        },
                        type: docxImageType,
                      }),
                    ],
                  }),
                ],
              }),
              new TableCell({
                width: {
                  size: 55,
                  type: WidthType.PERCENTAGE,
                },
                margins: {
                  left: 200, // Twips
                },
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
                          new TextRun({
                            text: `${key}. ${value}`,
                            size: 24, // 12pt
                          }),
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

  const handleDownloadDocx = async () => {
    if (!analysisResult || !imageUrl) return;
    try {
      const blob = await createDocxBlob([{ analysis: analysisResult, imageUrl }]);
      const url = window.URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = "exam-question.docx";
      a.click();
      window.URL.revokeObjectURL(url);
    } catch (e) {
      console.error("Error creating DOCX:", e);
      setError("Failed to create DOCX file.");
    }
  };

  const handleBatchDownloadDocx = async () => {
    if (batchResults.length === 0) return;
    try {
      const blob = await createDocxBlob(batchResults.map(r => ({ analysis: r.analysis, imageUrl: r.imageUrl })));
      const url = window.URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = "exam-questions-batch.docx";
      a.click();
      window.URL.revokeObjectURL(url);
    } catch (e) {
      console.error("Error creating Batch DOCX:", e);
      setError("Failed to create Batch DOCX file.");
    }
  };

  // Helper to determine badge color
  const getBadgeStyle = (type: string) => {
    switch(type) {
      case 'sign': return 'bg-red-100 text-red-700';
      case 'notice': return 'bg-blue-100 text-blue-700';
      case 'message': return 'bg-green-100 text-green-700';
      default: return 'bg-gray-100 text-gray-700';
    }
  };

  return (
    <div className="min-h-screen flex flex-col bg-gray-50 text-gray-800 font-sans">
      
      {/* --- HEADER --- */}
      <header className="bg-[#ff4500] text-white shadow-lg z-10">
        <div className="max-w-6xl mx-auto px-4 sm:px-6 py-4 sm:py-5 flex flex-col sm:flex-row justify-between items-center gap-4 sm:gap-0">
          <div className="text-center sm:text-left">
             <h1 className="text-xl sm:text-2xl font-bold tracking-tight">Sign Exam Generator</h1>
             <p className="text-[#ffe0cc] text-xs font-medium uppercase tracking-wider mt-1">Professional Exam Content Tool</p>
          </div>
          <div className="flex flex-col items-center sm:items-end">
            <span className="text-white/80 text-xs sm:text-sm italic">Powering Education with AI</span>
            <span className="text-white font-bold text-sm mt-1">Zalo 0913.885.221 (Ông Giáo)</span>
          </div>
        </div>
      </header>

      {/* --- BODY --- */}
      <main className="flex-grow max-w-5xl w-full mx-auto px-4 sm:px-6 py-6 sm:py-10 space-y-6 sm:space-y-8">

        {/* --- Instructions Section --- */}
        <div className="bg-white border border-gray-200 shadow-sm rounded-xl p-6 sm:p-8 mb-4">
           <h2 className="text-xl font-bold text-gray-800 mb-6 flex items-center gap-2">
             <svg xmlns="http://www.w3.org/2000/svg" className="h-6 w-6 text-[#ff4500]" fill="none" viewBox="0 0 24 24" stroke="currentColor">
               <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M13 16h-1v-4h-1m1-4h.01M21 12a9 9 0 11-18 0 9 9 0 0118 0z" />
             </svg>
             How it Works
           </h2>
           <div className="grid grid-cols-1 md:grid-cols-4 gap-8">
              {/* Step 1 */}
              <div className="flex flex-col items-center text-center relative">
                 <div className="w-14 h-14 rounded-full bg-[#fff5f0] text-[#ff4500] flex items-center justify-center mb-4 shadow-sm border border-[#ffe0cc]">
                   <svg xmlns="http://www.w3.org/2000/svg" className="h-7 w-7" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                     <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M9 19v-6a2 2 0 00-2-2H5a2 2 0 00-2 2v6a2 2 0 002 2h2a2 2 0 002-2zm0 0V9a2 2 0 012-2h2a2 2 0 012 2v10m-6 0a2 2 0 002 2h2a2 2 0 002-2m0 0V5a2 2 0 012-2h2a2 2 0 012 2v14a2 2 0 01-2 2h-2a2 2 0 01-2-2z" />
                   </svg>
                 </div>
                 <h4 className="font-bold text-gray-900 mb-1">1. Select Level</h4>
                 <p className="text-sm text-gray-500">Choose A2 or B1 to set the difficulty.</p>
                 <div className="hidden md:block absolute top-7 -right-4 w-8 border-t-2 border-dashed border-gray-200"></div>
              </div>
              
              {/* Step 2 */}
              <div className="flex flex-col items-center text-center relative">
                 <div className="w-14 h-14 rounded-full bg-[#fff5f0] text-[#ff4500] flex items-center justify-center mb-4 shadow-sm border border-[#ffe0cc]">
                   <svg xmlns="http://www.w3.org/2000/svg" className="h-7 w-7" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                     <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M11 5H6a2 2 0 00-2 2v11a2 2 0 002 2h11a2 2 0 002-2v-5m-1.414-9.414a2 2 0 112.828 2.828L11.828 15H9v-2.828l8.586-8.586z" />
                   </svg>
                 </div>
                 <h4 className="font-bold text-gray-900 mb-1">2. Input Data</h4>
                 <p className="text-sm text-gray-500">Type a question or upload a reference image.</p>
                 <div className="hidden md:block absolute top-7 -right-4 w-8 border-t-2 border-dashed border-gray-200"></div>
              </div>

              {/* Step 3 */}
              <div className="flex flex-col items-center text-center relative">
                 <div className="w-14 h-14 rounded-full bg-[#fff5f0] text-[#ff4500] flex items-center justify-center mb-4 shadow-sm border border-[#ffe0cc]">
                   <svg xmlns="http://www.w3.org/2000/svg" className="h-7 w-7" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                     <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M19.428 15.428a2 2 0 00-1.022-.547l-2.387-.477a6 6 0 00-3.86.517l-.318.158a6 6 0 01-3.86.517L6.05 15.21a2 2 0 00-1.806.547M8 4h8l-1 1v5.172a2 2 0 00.586 1.414l5 5c1.26 1.26.367 3.414-1.415 3.414H4.828c-1.782 0-2.674-2.154-1.414-3.414l5-5A2 2 0 009 10.172V5L8 4z" />
                   </svg>
                 </div>
                 <h4 className="font-bold text-gray-900 mb-1">3. Generate</h4>
                 <p className="text-sm text-gray-500">AI analyzes and generates content.</p>
                 <div className="hidden md:block absolute top-7 -right-4 w-8 border-t-2 border-dashed border-gray-200"></div>
              </div>

              {/* Step 4 */}
              <div className="flex flex-col items-center text-center">
                 <div className="w-14 h-14 rounded-full bg-[#fff5f0] text-[#ff4500] flex items-center justify-center mb-4 shadow-sm border border-[#ffe0cc]">
                   <svg xmlns="http://www.w3.org/2000/svg" className="h-7 w-7" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                     <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M4 16v1a3 3 0 003 3h10a3 3 0 003-3v-1m-4-4l-4 4m0 0l-4-4m4 4V4" />
                   </svg>
                 </div>
                 <h4 className="font-bold text-gray-900 mb-1">4. Download</h4>
                 <p className="text-sm text-gray-500">Get your exam questions in DOCX.</p>
              </div>
           </div>
        </div>

        {/* Global Controls: Difficulty Level */}
        <div className="bg-[#fff5f0] border border-gray-200 shadow-sm rounded-xl p-6 flex flex-col sm:flex-row sm:items-center justify-between gap-6 transition hover:shadow-md">
           <div>
             <h3 className="text-lg font-bold text-gray-800 flex items-center gap-2">
               <svg xmlns="http://www.w3.org/2000/svg" className="h-5 w-5 text-[#ff4500]" viewBox="0 0 20 20" fill="currentColor">
                 <path fillRule="evenodd" d="M12 7a1 1 0 110-2h5a1 1 0 011 1v5a1 1 0 11-2 0V8.414l-4.293 4.293a1 1 0 01-1.414 0L8 10.414l-4.293 4.293a1 1 0 01-1.414-1.414l5-5a1 1 0 011.414 0L11 10.586 14.586 7H12z" clipRule="evenodd" />
               </svg>
               Target Audience Level
             </h3>
             <p className="text-sm text-gray-500 mt-1">Select the CEFR proficiency level for generated questions.</p>
           </div>
           
           <div className="flex bg-gray-100 p-1 rounded-lg">
              {/* Level A2 */}
              <button
                onClick={() => setSelectedLevel("A2")}
                className={`px-5 py-2.5 rounded-md text-sm font-semibold transition-all duration-200 ${
                  selectedLevel === "A2"
                    ? "bg-white text-[#ff4500] shadow-sm ring-1 ring-gray-200"
                    : "text-gray-500 hover:text-gray-700"
                }`}
              >
                A2 (Elementary)
              </button>
              {/* Level B1 */}
              <button
                onClick={() => setSelectedLevel("B1")}
                className={`px-5 py-2.5 rounded-md text-sm font-semibold transition-all duration-200 ${
                  selectedLevel === "B1"
                    ? "bg-white text-[#ff4500] shadow-sm ring-1 ring-gray-200"
                    : "text-gray-500 hover:text-gray-700"
                }`}
              >
                B1 (Intermediate)
              </button>
           </div>
        </div>

        {/* Input Section */}
        <div className="bg-white rounded-xl shadow-lg border border-gray-100 overflow-hidden">
          
          {/* Tabs */}
          <div className="flex border-b border-gray-100">
            <button
              onClick={() => setActiveTab("manual")}
              className={`flex-1 py-4 text-sm transition-all duration-200 border-b-2 ${
                activeTab === "manual"
                  ? "border-[#ff4500] text-[#cf3700] bg-[#ffe0cc] font-bold shadow-[inset_0_-2px_4px_rgba(0,0,0,0.05)]"
                  : "border-transparent text-gray-500 hover:text-gray-700 hover:bg-gray-50 font-medium"
              }`}
            >
              Manual Input
            </button>
            <button
              onClick={() => setActiveTab("upload")}
              className={`flex-1 py-4 text-sm transition-all duration-200 border-b-2 ${
                activeTab === "upload"
                  ? "border-[#ff4500] text-[#cf3700] bg-[#ffe0cc] font-bold shadow-[inset_0_-2px_4px_rgba(0,0,0,0.05)]"
                  : "border-transparent text-gray-500 hover:text-gray-700 hover:bg-gray-50 font-medium"
              }`}
            >
              Upload Reference Image
            </button>
          </div>

          <div className="p-4 sm:p-8">
            {activeTab === "manual" ? (
              // MANUAL INPUT MODE
              <div className="animate-fade-in">
                <label htmlFor="questionInput" className="block text-sm font-semibold text-gray-700 mb-2">
                  Input Question Text
                </label>
                <textarea
                  id="questionInput"
                  className="w-full h-40 p-4 border border-gray-300 rounded-lg focus:ring-2 focus:ring-[#ff4500] focus:border-[#ff4500] outline-none transition text-base resize-none bg-gray-50 focus:bg-white"
                  placeholder="e.g., [NO ENTRY] You cannot enter this road. A. True B. False..."
                  value={inputText}
                  onChange={(e) => setInputText(e.target.value)}
                />
                <div className="mt-6 flex justify-end">
                  <button
                    onClick={handleGenerate}
                    disabled={loadingStep === "analyzing" || loadingStep === "drawing"}
                    className={`w-full sm:w-auto px-8 py-3 rounded-lg text-white font-bold tracking-wide shadow-md transition-transform active:scale-95 ${
                      loadingStep === "analyzing" || loadingStep === "drawing"
                        ? "bg-gray-400 cursor-not-allowed"
                        : "bg-[#ff4500] hover:bg-[#e03e00] hover:shadow-lg"
                    }`}
                  >
                    {loadingStep === "analyzing" && "Classifying & Analyzing..."}
                    {loadingStep === "drawing" && "Generating Illustration..."}
                    {loadingStep === "idle" || loadingStep === "done" ? "Generate Content" : ""}
                  </button>
                </div>
              </div>
            ) : (
              // UPLOAD MODE
              <div className="animate-fade-in space-y-6">
                <div className="border-2 border-dashed border-gray-300 rounded-lg p-6 sm:p-8 text-center hover:border-[#ff4500] hover:bg-[#fff5f0] transition-colors cursor-pointer relative group">
                  <input
                    type="file"
                    ref={fileInputRef}
                    accept="image/*"
                    onChange={handleReferenceUpload}
                    className="absolute inset-0 w-full h-full opacity-0 cursor-pointer"
                  />
                  <div className="pointer-events-none">
                     <svg xmlns="http://www.w3.org/2000/svg" className="h-12 w-12 mx-auto text-gray-400 group-hover:text-[#ff4500] transition-colors mb-4" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                       <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={1.5} d="M4 16l4.586-4.586a2 2 0 012.828 0L16 16m-2-2l1.586-1.586a2 2 0 012.828 0L20 14m-6-6h.01M6 20h12a2 2 0 002-2V6a2 2 0 00-2-2H6a2 2 0 00-2 2v12a2 2 0 002 2z" />
                     </svg>
                     <p className="text-gray-600 font-medium group-hover:text-[#ff4500]">Click to upload or drag and drop</p>
                     <p className="text-xs text-gray-400 mt-2">JPG, PNG (Max 5MB)</p>
                     <p className="text-sm text-[#ff4500] font-medium mt-4">Selected Level: {selectedLevel}</p>
                  </div>
                </div>

                {referenceImage && (
                  <div className="flex items-center gap-4 bg-gray-50 p-4 rounded-lg border border-gray-200">
                    <img src={referenceImage} alt="Reference" className="w-16 h-16 object-cover rounded border border-gray-300" />
                    <div>
                      <p className="text-sm font-semibold text-gray-800">Reference Image Loaded</p>
                      <p className="text-xs text-gray-500">Ready for analysis</p>
                    </div>
                  </div>
                )}

                {isAnalyzingRef && (
                  <div className="py-10 text-center text-gray-500">
                    <div className="inline-block w-8 h-8 border-4 border-[#ff4500] border-t-transparent rounded-full animate-spin mb-3"></div>
                    <p className="font-medium">Analyzing image structure...</p>
                    <p className="text-xs">Creating variations for Level {selectedLevel}</p>
                  </div>
                )}

                {suggestions.length > 0 && !isBatchProcessing && batchResults.length === 0 && (
                  <div className="bg-gray-50 rounded-lg p-6 border border-gray-200 mt-6">
                    <div className="flex justify-between items-center mb-4">
                      <h3 className="text-lg font-bold text-gray-800">Suggested Questions</h3>
                      <button 
                        onClick={handleSelectAll}
                        className="text-sm text-[#ff4500] hover:text-[#e03e00] font-semibold"
                      >
                         {selectedIndices.size === suggestions.length ? "Deselect All" : "Select All"}
                      </button>
                    </div>
                    
                    <div className="space-y-3">
                      {suggestions.map((suggestion, idx) => {
                        const isSelected = selectedIndices.has(idx);
                        return (
                          <div 
                            key={idx}
                            className={`p-4 rounded-lg border transition-all cursor-pointer flex items-start gap-4 ${
                              isSelected 
                                ? "border-[#ff4500] bg-white shadow-sm ring-1 ring-[#ff4500]" 
                                : "border-gray-200 bg-white hover:bg-gray-50"
                            }`}
                            onClick={() => toggleSelection(idx)}
                          >
                            <div className={`mt-1 flex-shrink-0 w-5 h-5 rounded border flex items-center justify-center ${
                                isSelected ? "bg-[#ff4500] border-[#ff4500]" : "border-gray-300 bg-white"
                            }`}>
                                {isSelected && <svg className="w-3 h-3 text-white" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth="3" d="M5 13l4 4L19 7"></path></svg>}
                            </div>
                            <div className="flex-1">
                               <pre className="text-sm text-gray-700 whitespace-pre-wrap font-sans leading-relaxed">{suggestion}</pre>
                               <div className="mt-2 text-right">
                                 <button 
                                  onClick={(e) => { e.stopPropagation(); handleUseSingleSuggestion(suggestion); }}
                                  className="text-xs font-semibold text-[#ff4500] hover:underline"
                                 >
                                   Edit Individually
                                 </button>
                               </div>
                            </div>
                          </div>
                        );
                      })}
                    </div>

                    <div className="mt-8 flex justify-end">
                      <button
                        onClick={handleBatchGenerate}
                        disabled={selectedIndices.size === 0}
                        className={`w-full sm:w-auto px-8 py-3 rounded-lg text-white font-bold shadow-md transition-all ${
                          selectedIndices.size === 0
                            ? "bg-gray-400 cursor-not-allowed"
                            : "bg-[#ff4500] hover:bg-[#e03e00] hover:shadow-lg active:scale-95"
                        }`}
                      >
                        Generate {selectedIndices.size} Selected Questions
                      </button>
                    </div>
                  </div>
                )}

                {/* Batch Processing Status */}
                {isBatchProcessing && (
                  <div className="py-12 text-center bg-gray-50 rounded-xl border border-dashed border-gray-300">
                    <div className="inline-block animate-spin rounded-full h-10 w-10 border-4 border-[#ff4500] border-t-transparent mb-4"></div>
                    <p className="text-xl font-bold text-gray-800">Generating Batch...</p>
                    <p className="text-gray-500 mt-2">{batchProgress}</p>
                  </div>
                )}

                {/* Batch Results */}
                {batchResults.length > 0 && (
                  <div className="mt-8 animate-fade-in">
                    <div className="flex flex-col sm:flex-row justify-between items-center mb-6 gap-4">
                      <h3 className="text-xl font-bold text-gray-800">Batch Results ({batchResults.length})</h3>
                      <button
                        onClick={handleBatchDownloadDocx}
                        className="w-full sm:w-auto flex items-center justify-center space-x-2 px-6 py-2.5 bg-green-600 text-white rounded-lg hover:bg-green-700 font-bold transition shadow-md hover:shadow-lg"
                      >
                        <svg xmlns="http://www.w3.org/2000/svg" className="h-5 w-5" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                          <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M4 16v1a3 3 0 003 3h10a3 3 0 003-3v-1m-4-4l-4 4m0 0l-4-4m4 4V4" />
                        </svg>
                        <span>Download All (DOCX)</span>
                      </button>
                    </div>
                    
                    <div className="grid grid-cols-1 gap-8">
                      {batchResults.map((result, i) => (
                        <div key={i} className="bg-white p-6 border border-gray-200 rounded-xl shadow-md hover:shadow-lg transition-shadow">
                          <div className="flex flex-col md:flex-row gap-8">
                            <div className="w-full md:w-1/3 flex-shrink-0">
                              <div className="bg-gray-100 rounded-lg border border-gray-200 aspect-video flex items-center justify-center p-4">
                                 <img src={result.imageUrl} alt={`Result ${i+1}`} className="w-full h-full object-contain drop-shadow-sm" />
                              </div>
                            </div>
                            <div className="w-full md:w-2/3">
                              <div className="flex items-center gap-3 mb-3">
                                <span className="bg-gray-800 text-white text-xs font-bold px-2 py-1 rounded">Q{i+1}</span>
                                <span className={`text-xs font-bold uppercase tracking-wider px-2 py-1 rounded ${getBadgeStyle(result.analysis.type)}`}>
                                  {result.analysis.type}
                                </span>
                              </div>
                              <p className="font-bold text-lg text-gray-900 mb-4">{result.analysis.docx_content.question}</p>
                              <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
                                 <div className="bg-gray-50 p-3 rounded border border-gray-100 text-sm"><span className="font-bold text-[#ff4500]">A.</span> {result.analysis.docx_content.options.A}</div>
                                 <div className="bg-gray-50 p-3 rounded border border-gray-100 text-sm"><span className="font-bold text-[#ff4500]">B.</span> {result.analysis.docx_content.options.B}</div>
                                 <div className="bg-gray-50 p-3 rounded border border-gray-100 text-sm"><span className="font-bold text-[#ff4500]">C.</span> {result.analysis.docx_content.options.C}</div>
                                 <div className="bg-gray-50 p-3 rounded border border-gray-100 text-sm"><span className="font-bold text-[#ff4500]">D.</span> {result.analysis.docx_content.options.D}</div>
                              </div>
                            </div>
                          </div>
                        </div>
                      ))}
                    </div>
                    
                    <div className="mt-10 flex justify-center">
                      <button 
                        onClick={() => { setBatchResults([]); setSelectedIndices(new Set()); setSuggestions([]); setReferenceImage(null); }}
                        className="text-gray-500 hover:text-[#ff4500] font-medium underline underline-offset-4 transition-colors"
                      >
                        Start Over
                      </button>
                    </div>
                  </div>
                )}
              </div>
            )}
          </div>
        </div>

        {/* Error Display */}
        {error && (
          <div className="bg-red-50 text-red-700 p-4 rounded-lg border border-red-200 flex items-start gap-3 shadow-sm">
            <svg xmlns="http://www.w3.org/2000/svg" className="h-6 w-6 flex-shrink-0" fill="none" viewBox="0 0 24 24" stroke="currentColor">
              <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M12 9v2m0 4h.01m-6.938 4h13.856c1.54 0 2.502-1.667 1.732-3L13.732 4c-.77-1.333-2.694-1.333-3.464 0L3.34 16c-.77 1.333.192 3 1.732 3z" />
            </svg>
            <div>
              <strong className="font-bold block">Error Encountered</strong>
              <span className="text-sm">{error}</span>
            </div>
          </div>
        )}

        {/* Results Section (Single Manual) */}
        {analysisResult && activeTab === "manual" && (
          <section className="bg-white p-4 sm:p-8 rounded-xl shadow-lg border border-gray-100 animate-slide-up">
            <div className="flex flex-col sm:flex-row justify-between items-center mb-6 sm:mb-8 pb-4 border-b border-gray-100 gap-4">
               <div className="flex items-center gap-4 w-full sm:w-auto">
                 <h2 className="text-xl sm:text-2xl font-bold text-gray-800">Generated Preview</h2>
                 <span className={`px-3 py-1 text-xs font-bold uppercase tracking-wider rounded-full ${getBadgeStyle(analysisResult.type)}`}>
                   Detected: {analysisResult.type}
                 </span>
               </div>
               {imageUrl && (
                 <button
                   onClick={handleDownloadDocx}
                   className="w-full sm:w-auto flex items-center justify-center space-x-2 px-6 py-2 border-2 border-gray-200 rounded-lg hover:border-[#ff4500] hover:text-[#ff4500] text-gray-600 font-bold transition-colors"
                 >
                   <svg xmlns="http://www.w3.org/2000/svg" className="h-5 w-5" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                     <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M4 16v1a3 3 0 003 3h10a3 3 0 003-3v-1m-4-4l-4 4m0 0l-4-4m4 4V4" />
                   </svg>
                   <span>Download DOCX</span>
                 </button>
               )}
            </div>

            <div className="flex flex-col lg:flex-row gap-10">
              {/* Left Column: Image */}
              <div className="w-full lg:w-5/12">
                <div className="text-xs font-bold text-gray-500 uppercase tracking-wider mb-2">Illustration</div>
                <div className={`w-full bg-white rounded-lg border-2 border-gray-100 shadow-inner flex items-center justify-center overflow-hidden relative ${
                    (analysisResult.type === "notice" || analysisResult.type === "message") ? "aspect-video" : "aspect-square"
                }`}>
                   {imageUrl ? (
                     <img src={imageUrl} alt="Generated Sign" className="w-full h-full object-contain p-2" />
                   ) : (
                     <div className="flex flex-col items-center text-gray-300 animate-pulse">
                        <svg xmlns="http://www.w3.org/2000/svg" className="h-12 w-12 mb-2" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                          <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M4 16l4.586-4.586a2 2 0 012.828 0L16 16m-2-2l1.586-1.586a2 2 0 012.828 0L20 14m-6-6h.01M6 20h12a2 2 0 002-2V6a2 2 0 00-2-2H6a2 2 0 00-2 2v12a2 2 0 002 2z" />
                        </svg>
                     </div>
                   )}
                </div>
                <div className="mt-4 p-4 bg-gray-50 text-gray-600 text-sm rounded-lg border border-gray-200 italic">
                  "{analysisResult.image_description}"
                </div>
              </div>

              {/* Right Column: Text */}
              <div className="w-full lg:w-7/12">
                 <div className="text-xs font-bold text-gray-500 uppercase tracking-wider mb-2">Exam Content</div>
                 <div className="p-8 border border-gray-200 rounded-xl bg-gray-50/50 h-full flex flex-col justify-center">
                    <p className="text-xl font-bold text-gray-900 mb-8 leading-snug">{analysisResult.docx_content.question}</p>
                    <div className="space-y-4">
                      {Object.entries(analysisResult.docx_content.options).map(([key, val]) => (
                        <div key={key} className="flex items-start group">
                          <span className="flex-shrink-0 w-8 h-8 rounded-full bg-[#ff4500] text-white flex items-center justify-center font-bold text-sm mr-4 shadow-sm group-hover:bg-[#e03e00] transition-colors">{key}</span>
                          <span className="text-lg text-gray-700 pt-0.5">{val}</span>
                        </div>
                      ))}
                    </div>
                 </div>
              </div>
            </div>
          </section>
        )}
      </main>

      {/* --- FOOTER --- */}
      <footer className="bg-gray-800 text-gray-400 py-8 border-t border-gray-700 mt-auto">
        <div className="container mx-auto px-4 text-center">
           <p className="mb-2 font-medium text-gray-300">English Exam Content Generator</p>
           <p className="text-sm">
             &copy; {new Date().getFullYear()} All rights reserved. <a href="https://globalsuccessfiles.com" target="_blank" rel="noopener noreferrer" className="hover:text-white underline decoration-dotted underline-offset-2 transition-colors">Designed by GlobalSuccessFiles.Com</a>
           </p>
        </div>
      </footer>
    </div>
  );
};

const root = createRoot(document.getElementById("root")!);
root.render(<App />);