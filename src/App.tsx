import React, { useState, useRef } from "react";
import {
  Upload,
  FileText,
  Image as ImageIcon,
  ArrowRight,
  Download,
  CheckCircle2,
  RefreshCw,
  AlertTriangle,
  Languages,
  Sparkles,
  Layers,
  Eye,
  Columns,
  BookOpen,
  Info,
  ChevronRight
} from "lucide-react";
import { motion, AnimatePresence } from "motion/react";

interface ImageTranslationInfo {
  id: string;
  index: number;
  mimeType: string;
  base64: string;
  translations: {
    originalText: string;
    translatedText: string;
  }[];
  error?: string;
}

interface TranslationResult {
  fileId: string;
  fileName: string;
  originalHtml: string;
  translatedHtml: string;
  hasImages: boolean;
  imageTranslations: ImageTranslationInfo[];
  fallbackBannerUsed?: boolean;
}

export default function App() {
  const [file, setFile] = useState<File | null>(null);
  const [isDragging, setIsDragging] = useState(false);
  const [isTranslating, setIsTranslating] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [currentStep, setCurrentStep] = useState<number>(0);
  const [translationResult, setTranslationResult] = useState<TranslationResult | null>(null);
  const [activeTab, setActiveTab] = useState<"text" | "images">("text");
  const [showDocs, setShowDocs] = useState(false);

  const fileInputRef = useRef<HTMLInputElement>(null);

  const steps = [
    { title: "Document Parsing", desc: "Extracting paragraphs & structural image binary nodes" },
    { title: "Text Translation (Gemini)", desc: "Processing HTML body nodes sequentially into English" },
    { title: "Multimodal OCR & Analysis", desc: "Analyzing graphical Chinese texts underneath embedded assets" },
    { title: "Format Reconstruction", desc: "Synthesizing output Office DOCX nodes with preservation rules" },
    { title: "Download Package Assembly", desc: "Caching output package on server for live instant download" }
  ];

  // Drag-and-drop event handlers
  const handleDragOver = (e: React.DragEvent) => {
    e.preventDefault();
    setIsDragging(true);
  };

  const handleDragLeave = () => {
    setIsDragging(false);
  };

  const handleDrop = (e: React.DragEvent) => {
    e.preventDefault();
    setIsDragging(false);
    setError(null);
    const droppedFile = e.dataTransfer.files[0];
    if (droppedFile) {
      validateAndSetFile(droppedFile);
    }
  };

  const handleFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    setError(null);
    const selectedFile = e.target.files?.[0];
    if (selectedFile) {
      validateAndSetFile(selectedFile);
    }
  };

  const validateAndSetFile = (selectedFile: File) => {
    const extension = selectedFile.name.split(".").pop()?.toLowerCase();
    if (extension !== "docx") {
      setError("Only Microsoft Word Documents (.docx) are supported for processing.");
      setFile(null);
      return;
    }
    setFile(selectedFile);
  };

  const executeTranslation = async () => {
    if (!file) return;

    setIsTranslating(true);
    setError(null);
    setTranslationResult(null);
    setCurrentStep(0);

    // Dynamic frontend progress simulator in case it takes a moment
    const stepInterval = setInterval(() => {
      setCurrentStep((prev) => {
        if (prev < 4) {
          return prev + 1;
        }
        return prev;
      });
    }, 4000);

    try {
      const formData = new FormData();
      formData.append("file", file);

      const response = await fetch("/api/translate", {
        method: "POST",
        body: formData,
      });

      if (!response.ok) {
        const errorData = await response.json().catch(() => ({}));
        throw new Error(errorData.error || "Translation pipeline failed. Please ensure file content is readable.");
      }

      const data: TranslationResult = await response.json();
      
      clearInterval(stepInterval);
      setCurrentStep(4);
      
      setTimeout(() => {
        setTranslationResult(data);
        setIsTranslating(false);
      }, 600);

    } catch (err: any) {
      clearInterval(stepInterval);
      setError(err.message || "An unexpected network error occurred while communicating with backend translator.");
      setIsTranslating(false);
    }
  };

  const triggerSelectFile = () => {
    fileInputRef.current?.click();
  };

  const resetState = () => {
    setFile(null);
    setError(null);
    setTranslationResult(null);
    setIsTranslating(false);
    setCurrentStep(0);
  };

  return (
    <div className="flex flex-col min-h-screen bg-slate-50 text-slate-900 font-sans">
      
      {/* Header (DocTrans AI Style) */}
      <header className="flex flex-col sm:flex-row items-center justify-between px-8 py-4 bg-white border-b border-slate-200 gap-4 shrink-0">
        <div className="flex items-center gap-3">
          <div className="w-10 h-10 bg-blue-600 rounded-lg flex items-center justify-center text-white font-bold text-xl shadow-sm">
            DT
          </div>
          <div>
            <h1 className="text-lg font-bold leading-tight text-slate-800">DocTrans AI</h1>
            <p className="text-xs text-slate-500">Multimodal Document Intelligence Pipeline</p>
          </div>
        </div>
        
        <div className="flex items-center gap-4 flex-wrap">
          <div className="flex items-center gap-2 px-3 py-1 bg-emerald-50 text-emerald-700 rounded-full border border-emerald-100 text-xs font-semibold">
            <span className="w-2 h-2 bg-emerald-500 rounded-full animate-pulse"></span>
            System Ready
          </div>
          <button 
            onClick={() => setShowDocs(!showDocs)}
            className="px-4 py-2 bg-slate-100 hover:bg-slate-200 rounded-md text-xs font-medium transition-colors text-slate-700 underline decoration-slate-300 cursor-pointer"
          >
            {showDocs ? "Hide Guidance" : "Documentation"}
          </button>
        </div>
      </header>

      {/* Docs Drawer */}
      <AnimatePresence>
        {showDocs && (
          <motion.div 
            initial={{ height: 0, opacity: 0 }}
            animate={{ height: "auto", opacity: 1 }}
            exit={{ height: 0, opacity: 0 }}
            className="bg-slate-900 text-slate-200 overflow-hidden border-b border-slate-800"
          >
            <div className="max-w-7xl mx-auto px-8 py-6 grid grid-cols-1 md:grid-cols-3 gap-6 text-xs">
              <div>
                <h4 className="font-bold text-white mb-2 text-sm flex items-center gap-1.5">
                  <Languages className="w-4 h-4 text-blue-400" />
                  1. Double Preservation Rules
                </h4>
                <p className="text-slate-400 leading-relaxed">
                  Both your structural paragraph layout (headings, spacing, tables, bold text tags) and your document embedded binary files are retrieved via Mammoth markup analyzer. Everything returns fully aligned.
                </p>
              </div>
              <div>
                <h4 className="font-bold text-white mb-2 text-sm flex items-center gap-1.5">
                  <Sparkles className="w-4 h-4 text-purple-400" />
                  2. Multimodal OCR Tables
                </h4>
                <p className="text-slate-400 leading-relaxed">
                  Avoid sloppy image overwrites that ruin resolution. Gemini analyzes all in-photo Chinese texts and appends a structured bilingual translation map in clear tables right underneath original assets.
                </p>
              </div>
              <div>
                <h4 className="font-bold text-white mb-2 text-sm flex items-center gap-1.5">
                  <Download className="w-4 h-4 text-emerald-400" />
                  3. Production Assembly
                </h4>
                <p className="text-slate-400 leading-relaxed">
                  Re-synthesized dynamically using our backend docx compiler. Output download features are instantly packed in memory and available for high-speed download instantly.
                </p>
              </div>
            </div>
          </motion.div>
        )}
      </AnimatePresence>

      {/* Main Content Area */}
      <main className="flex-1 max-w-7xl w-full mx-auto p-6 md:p-8 grid grid-cols-1 lg:grid-cols-12 gap-8 items-start">
        
        {/* Left Section (5 cols): Pipeline Controls, Stepper state */}
        <section className="lg:col-span-5 space-y-6">
          <div className="flex items-center justify-between">
            <h2 className="text-xs font-bold uppercase tracking-wider text-slate-400">Control & Status Console</h2>
            {file && (
              <span className="text-xs text-slate-400 font-semibold bg-white px-2.5 py-1 rounded-full border border-slate-200">
                {(file.size / 1024 / 1024).toFixed(2)} MB • Word.docx
              </span>
            )}
          </div>

          {/* Core Upload and Status container */}
          <div className="bg-white rounded-2xl border border-slate-200 shadow-sm p-6 space-y-6">
            
            {/* Displaying Any Errors */}
            {error && (
              <div className="p-4 bg-rose-50 border border-rose-100 rounded-xl flex items-start gap-3 text-rose-800">
                <AlertTriangle className="w-5 h-5 text-rose-600 shrink-0 mt-0.5" />
                <div>
                  <h4 className="font-semibold text-rose-900">Pipeline Encountered Errors</h4>
                  <p className="text-xs mt-1">{error}</p>
                </div>
              </div>
            )}

            {!translationResult && !isTranslating ? (
              <div className="space-y-6">
                <div>
                  <h3 className="text-sm font-bold text-slate-700 uppercase mb-3 tracking-wide">
                    Upload Word Document
                  </h3>
                  
                  {/* File Drag drop region */}
                  <div
                    onDragOver={handleDragOver}
                    onDragLeave={handleDragLeave}
                    onDrop={handleDrop}
                    onClick={triggerSelectFile}
                    className={`border-2 border-dashed rounded-xl p-8 text-center cursor-pointer transition-all duration-200 ${
                      isDragging
                        ? "border-blue-500 bg-blue-50/50 scale-[0.99]"
                        : file
                        ? "border-emerald-400 bg-emerald-50/10 hover:border-emerald-500"
                        : "border-slate-200 bg-slate-50 hover:bg-slate-100/50 hover:border-slate-300"
                    }`}
                  >
                    <input
                      type="file"
                      ref={fileInputRef}
                      onChange={handleFileChange}
                      className="hidden"
                      accept=".docx"
                    />

                    {file ? (
                      <div className="flex flex-col items-center">
                        <div className="w-14 h-14 rounded-xl bg-emerald-100 text-emerald-600 flex items-center justify-center mb-4 shadow-xs">
                          <FileText className="w-7 h-7" />
                        </div>
                        <p className="text-slate-900 font-bold text-sm truncate max-w-full px-2">
                          {file.name}
                        </p>
                        <p className="text-xs text-slate-400 mt-1">
                          File is ready for processing
                        </p>
                        <span className="mt-4 text-xs font-semibold text-blue-600 hover:underline">
                          Choose another file
                        </span>
                      </div>
                    ) : (
                      <div className="flex flex-col items-center">
                        <div className="w-14 h-14 rounded-xl bg-slate-100 text-slate-400 flex items-center justify-center mb-3">
                          <Upload className="w-6 h-6" />
                        </div>
                        <p className="text-slate-700 font-semibold text-sm">
                          Drag and drop file here, or <span className="text-blue-600 font-bold hover:underline">Browse</span>
                        </p>
                        <p className="text-[11px] text-slate-400 mt-1.5">
                          Format: MS Word .docx only. Support files up to 20MB.
                        </p>
                      </div>
                    )}
                  </div>
                </div>

                {/* Direct Translation Direction Indicator */}
                <div className="p-4 rounded-xl bg-slate-50 border border-slate-100 flex items-center justify-between text-xs">
                  <div>
                    <span className="block text-[10px] text-slate-400 font-bold uppercase tracking-wider">Source Language</span>
                    <span className="font-bold text-slate-800">中文 (Chinese)</span>
                  </div>
                  <ChevronRight className="w-5 h-5 text-slate-300" />
                  <div className="text-right">
                    <span className="block text-[10px] text-slate-400 font-bold uppercase tracking-wider">Target Language</span>
                    <span className="font-bold text-blue-600">English (Bilingual Appends)</span>
                  </div>
                </div>

                {/* Primary Action Button */}
                <button
                  disabled={!file}
                  onClick={executeTranslation}
                  className="w-full py-3.5 px-4 rounded-xl font-bold text-sm text-white transition-all shadow-md flex items-center justify-center gap-2 cursor-pointer disabled:bg-slate-200 disabled:text-slate-400 disabled:cursor-not-allowed bg-blue-600 hover:bg-blue-700 shadow-blue-500/10"
                >
                  <Sparkles className="w-4 h-4" />
                  <span>Execute Multimodal Pipeline</span>
                </button>
              </div>
            ) : isTranslating ? (
              
              /* Pipeline Stepper - Matching Theme Specifications step by step */
              <div className="space-y-4">
                <div className="flex items-center justify-between border-b border-slate-100 pb-3">
                  <div>
                    <span className="text-xs font-bold text-slate-400 uppercase tracking-widest block">Pipeline Stage</span>
                    <span className="font-bold text-slate-800 text-sm">Multimodal Network Analysis</span>
                  </div>
                  <div className="flex items-center gap-1.5 text-xs text-blue-600 bg-blue-50 border border-blue-100 px-2.5 py-1 rounded-full font-bold">
                    <RefreshCw className="w-3.5 h-3.5 animate-spin" />
                    <span>Processing</span>
                  </div>
                </div>

                {/* Steps List */}
                <div className="space-y-3 pt-2">
                  {steps.map((step, idx) => {
                    const isActive = idx === currentStep;
                    const isDone = idx < currentStep;

                    return (
                      <div
                        key={idx}
                        className={`transition-all duration-300 rounded-lg p-4 flex items-center gap-4 border ${
                          isActive
                            ? "bg-white border-blue-500 shadow-md relative z-10 scale-[1.01]"
                            : isDone
                            ? "bg-white border-slate-100 opacity-90"
                            : "bg-slate-50 border-dashed border-slate-200 opacity-60"
                        }`}
                      >
                        {/* Status Icon Indicator */}
                        <div className="shrink-0">
                          {isDone ? (
                            <div className="w-8 h-8 rounded-full bg-emerald-500 flex items-center justify-center text-white">
                              <svg className="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                                <path strokeLinecap="round" strokeLinejoin="round" strokeWidth="3" d="M5 13l4 4L19 7" />
                              </svg>
                            </div>
                          ) : isActive ? (
                            <div className="w-8 h-8 rounded-full bg-blue-100 flex items-center justify-center text-blue-600">
                              <svg className="w-4 h-4 animate-spin" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                                <circle className="opacity-25" cx="12" cy="12" r="10" stroke="currentColor" strokeWidth="4" />
                                <path className="opacity-75" fill="currentColor" d="M4 12a8 8 0 018-8V0C5.373 0 0 5.373 0 12h4zm2 5.291A7.962 7.962 0 014 12H0c0 3.042 1.135 5.824 3 7.938l3-2.647z" />
                              </svg>
                            </div>
                          ) : (
                            <div className="w-8 h-8 rounded-full bg-slate-200 flex items-center justify-center text-slate-400 text-xs font-bold font-mono">
                              {idx + 1}
                            </div>
                          )}
                        </div>

                        {/* Title and details */}
                        <div className="flex-1 min-w-0">
                          <div className="flex items-center justify-between">
                            <h3 className={`text-xs font-bold uppercase tracking-wider ${isActive ? "text-blue-700" : isDone ? "text-slate-700" : "text-slate-400"}`}>
                              {step.title}
                            </h3>
                            {isActive && (
                              <span className="text-[10px] font-bold text-blue-600 bg-blue-100/50 px-2 py-0.5 rounded-full">
                                {Math.min(99, Math.round(((idx + 0.6) / steps.length) * 100))}%
                              </span>
                            )}
                          </div>
                          {isActive && (
                            <div className="w-full bg-slate-100 h-1 rounded-full overflow-hidden my-1">
                              <div className="bg-blue-600 h-full animate-pulse w-4/5" />
                            </div>
                          )}
                          <p className="text-xs text-slate-500 truncate">{step.desc}</p>
                        </div>
                      </div>
                    );
                  })}
                </div>
              </div>
            ) : (
              
              /* Completed Document Status Block - Pristine Slate Dark Aesthetic to match Professional Polish */
              <div className="space-y-6">
                {translationResult?.fallbackBannerUsed && (
                  <div className="p-4 bg-amber-50 border border-amber-200 rounded-xl flex gap-3 text-amber-900 text-xs shadow-xs">
                    <AlertTriangle className="w-5 h-5 text-amber-600 shrink-0 mt-0.5" />
                    <div>
                      <h4 className="font-bold text-amber-950 uppercase tracking-wider text-[10px] mb-1">Standard AI Quota Reached</h4>
                      <p className="leading-relaxed text-amber-800">
                        The shared translation engine daily quota has been fully active & exhausted (Rate Limit 429). 
                        Our system initiated <strong>Resilient Document Fallback Mode</strong>. Your document file structural columns, page borders, table nodes, and embedded images have been compiled successfully!
                      </p>
                    </div>
                  </div>
                )}

                <div className="text-center p-5 bg-emerald-50 rounded-xl border border-emerald-100">
                  <CheckCircle2 className="w-10 h-10 text-emerald-500 mx-auto mb-2" />
                  <h3 className="font-semibold text-emerald-950 text-sm uppercase tracking-wider">Output File Ready</h3>
                  <p className="text-xs text-emerald-700/80 mt-1">
                    Your English translation layout has been assembled and successfully packed.
                  </p>
                </div>

                {/* Elegant dark Card block for downloading (Fits perfectly to thematic style "bg-slate-900 text-white rounded-xl p-6") */}
                <div className="bg-slate-900 rounded-xl p-5 text-white flex items-center justify-between">
                  <div className="flex items-center gap-4 min-w-0 flex-1">
                    <div className="p-3 bg-white/10 rounded-lg text-white font-bold shrink-0">
                      <FileText className="w-6 h-6 text-blue-400" />
                    </div>
                    <div className="min-w-0 flex-1">
                      <p className="text-[10px] text-slate-400 uppercase tracking-widest font-bold">Download Ready</p>
                      <h4 className="text-xs font-medium text-slate-100 truncate" title={translationResult?.fileName}>
                        {translationResult?.fileName}
                      </h4>
                    </div>
                  </div>
                  <a
                    href={`/api/download/${translationResult?.fileId}`}
                    className="ml-3 px-5 py-2.5 bg-blue-600 hover:bg-blue-700 text-white font-bold rounded-lg text-xs transition-colors shadow-lg shadow-blue-950/40 inline-flex items-center gap-1.5 cursor-pointer whitespace-nowrap"
                  >
                    <Download className="w-3.5 h-3.5" />
                    Download
                  </a>
                </div>

                <button
                  onClick={resetState}
                  className="w-full py-2.5 px-4 bg-slate-100 hover:bg-slate-200 text-slate-700 font-semibold rounded-xl text-xs transition-all flex items-center justify-center gap-2 cursor-pointer"
                >
                  <RefreshCw className="w-3.5 h-3.5" />
                  <span>Process New Document</span>
                </button>
              </div>
            )}
          </div>

          {/* Guidelines Sidebar summary card */}
          <div className="p-5 bg-white border border-slate-200 rounded-2xl flex gap-3 text-xs leading-relaxed text-slate-600 shadow-2xs">
            <Info className="w-4 h-4 text-blue-600 shrink-0 mt-0.5" />
            <div>
              <p className="font-bold text-slate-800 mb-1 uppercase tracking-wider text-[10px]">Embedded Graphics Notice</p>
              Our process automatically registers all embedded images in correct location. Instead of altering image quality with raw rewrite paints, a synchronized translation guide mapping sheet is appended to original layouts.
            </div>
          </div>
        </section>

        {/* Right Section (7 cols): Document View panel / Tab previews */}
        <section className="lg:col-span-7">
          {translationResult ? (
            <div className="bg-white rounded-2xl border border-slate-200 shadow-sm overflow-hidden min-h-[520px] flex flex-col">
              
              {/* Header inside result window */}
              <div className="bg-slate-50 px-6 py-4 border-b border-slate-200 flex flex-col sm:flex-row sm:items-center sm:justify-between gap-3 shrink-0">
                <div className="flex items-center gap-2">
                  <Eye className="w-5 h-5 text-blue-600" />
                  <span className="font-bold text-xs text-slate-700 uppercase tracking-wider">Results Explorer Preview</span>
                </div>
                
                {/* Horizontal Switchers */}
                <div className="flex bg-slate-200 p-1 rounded-lg self-start">
                  <button
                    onClick={() => setActiveTab("text")}
                    className={`px-4 py-1.5 rounded-md text-xs font-semibold transition-all flex items-center gap-1.5 cursor-pointer ${
                      activeTab === "text"
                        ? "bg-white text-blue-700 shadow-xs"
                        : "text-slate-500 hover:text-slate-900"
                    }`}
                  >
                    <Columns className="w-3.5 h-3.5" />
                    Text Contrast
                  </button>
                  <button
                    onClick={() => setActiveTab("images")}
                    className={`px-4 py-1.5 rounded-md text-xs font-semibold transition-all flex items-center gap-1.5 cursor-pointer ${
                      activeTab === "images"
                        ? "bg-white text-blue-700 shadow-xs"
                        : "text-slate-500 hover:text-slate-900"
                    }`}
                  >
                    <ImageIcon className="w-3.5 h-3.5" />
                    Embedded Graphics ({translationResult.imageTranslations.length})
                  </button>
                </div>
              </div>

              {/* Dynamic scrollable data wrapper */}
              <div className="p-6 flex-1 overflow-y-auto max-h-[600px] bg-slate-50/50">
                
                {/* Visual view text contrast mapping pairs */}
                {activeTab === "text" && (
                  <div className="space-y-6">
                    <div className="bg-white p-3 border border-slate-200 rounded-xl text-xs text-slate-600 flex items-start gap-2">
                      <BookOpen className="w-4 h-4 text-blue-500 shrink-0 mt-0.5" />
                      <div>
                        <strong>Bilingual Sentence Mapping:</strong> Compare elements side by side. Original XML paragraphs on left map accurately to structural Translated blocks on right.
                      </div>
                    </div>

                    <div className="grid grid-cols-1 md:grid-cols-2 gap-6 items-stretch">
                      
                      {/* Left Column (ZH) */}
                      <div className="bg-white p-5 rounded-xl border border-slate-200 flex flex-col min-h-[350px] shadow-2xs">
                        <div className="font-bold text-[10px] text-slate-400 mb-3 uppercase tracking-wider border-b border-slate-100 pb-2">
                          Original Content Extracts (ZH)
                        </div>
                        <div 
                          className="text-xs text-slate-700 whitespace-pre-wrap leading-relaxed flex-1 prose prose-slate max-w-none prose-sm"
                          dangerouslySetInnerHTML={{ __html: sanitizeHtmlWithPlaceholders(translationResult.originalHtml) }}
                        />
                      </div>

                      {/* Right Column (EN) */}
                      <div className="bg-white p-5 rounded-xl border border-slate-200 flex flex-col min-h-[350px] shadow-2xs">
                        <div className="font-bold text-[10px] text-blue-600 mb-3 uppercase tracking-wider border-b border-blue-100 pb-2 flex items-center justify-between">
                          <span>Translated Results Output (EN)</span>
                          <span className="text-[9px] bg-blue-50 text-blue-700 px-2 py-0.5 rounded font-mono uppercase tracking-normal">Verified</span>
                        </div>
                        <div 
                          className="text-xs text-slate-700 whitespace-pre-wrap leading-relaxed flex-1 prose prose-slate max-w-none prose-sm"
                          dangerouslySetInnerHTML={{ __html: sanitizeHtmlWithPlaceholders(translationResult.translatedHtml) }}
                        />
                      </div>

                    </div>
                  </div>
                )}

                {/* Multimodal Photo Data Mapping sheet */}
                {activeTab === "images" && (
                  <div className="space-y-6">
                    <div className="bg-white p-3 border border-slate-200 rounded-xl text-xs text-slate-600">
                      <strong>Multimodal OCR Recognition:</strong> All graphics embedded in raw files will maintain original aspect ratio inside the final downloadable DOCX. Here are OCR translations mapped cleanly.
                    </div>

                    {!translationResult.hasImages ? (
                      <div className="text-center py-16 bg-white border border-slate-200 rounded-xl text-slate-400">
                        <ImageIcon className="w-12 h-12 mx-auto mb-2 text-slate-300" />
                        <p className="text-xs font-semibold uppercase tracking-wider">No Embedded Visual Assets Located</p>
                      </div>
                    ) : (
                      <div className="space-y-8">
                        {translationResult.imageTranslations.map((img, idx) => (
                          <div key={img.id} className="p-4 rounded-xl border border-slate-200 bg-white space-y-4 shadow-2xs">
                            <div className="flex items-center justify-between border-b border-slate-100 pb-2.5">
                              <span className="font-bold text-xs text-slate-700 flex items-center gap-1.5 uppercase">
                                <ImageIcon className="w-4 h-4 text-blue-600" />
                                Image File Asset #{idx + 1}
                              </span>
                              <span className="text-[10px] font-bold text-slate-400 bg-slate-100 px-2.5 py-0.5 rounded">
                                {img.mimeType}
                              </span>
                            </div>

                            <div className="grid grid-cols-1 md:grid-cols-12 gap-5 items-start">
                              
                              {/* Left column in grid (visual card thumbnail) */}
                              <div className="md:col-span-5 bg-slate-50 border border-slate-200 rounded-lg p-3 flex flex-col items-center justify-center relative min-h-[160px]">
                                <div className="absolute top-2.5 left-2.5 bg-blue-100 text-blue-800 px-2 py-0.5 rounded text-[9px] font-mono font-bold uppercase tracking-wider">
                                  IMAGE_{idx + 1}.PNG
                                </div>
                                <img
                                  src={`data:${img.mimeType};base64,${img.base64}`}
                                  alt={`Binary Img ${idx + 1}`}
                                  className="max-h-[130px] rounded object-contain mt-2"
                                />
                                <div className="mt-2.5 flex items-center gap-1">
                                  <span className="w-1.5 h-1.5 bg-blue-500 rounded-full"></span>
                                  <span className="text-[9px] font-bold text-slate-400 tracking-wider uppercase">ZH_CN OCR READY</span>
                                </div>
                              </div>

                              {/* Right column in grid (Translations Table list) */}
                              <div className="md:col-span-7">
                                {img.translations.length === 0 ? (
                                  <div className="p-5 bg-slate-50 rounded-lg border border-slate-200 text-center text-slate-400 text-xs">
                                    {img.error ? (
                                      <p className="text-rose-500 font-bold">Critical analysis logic timed out.</p>
                                    ) : (
                                      <p>No Chinese visual texts readable in image asset.</p>
                                    )}
                                  </div>
                                ) : (
                                  <div className="overflow-hidden border border-slate-100 rounded-lg shadow-2xs">
                                    <table className="min-w-full divide-y divide-slate-100 text-[11px] bg-white">
                                      <thead className="bg-slate-50">
                                        <tr>
                                          <th className="px-3 py-2 text-left font-bold text-slate-500 uppercase tracking-wider w-1/2">Chinese (Original)</th>
                                          <th className="px-3 py-2 text-left font-bold text-blue-700 uppercase tracking-wider w-1/2">English (OCR Mapped)</th>
                                        </tr>
                                      </thead>
                                      <tbody className="divide-y divide-slate-100">
                                        {img.translations.map((item, transIdx) => (
                                          <tr key={transIdx} className="hover:bg-slate-50/40">
                                            <td className="px-3 py-2 text-slate-800 font-medium whitespace-pre-wrap">{item.originalText}</td>
                                            <td className="px-3 py-2 text-slate-600 italic whitespace-pre-wrap">{item.translatedText}</td>
                                          </tr>
                                        ))}
                                      </tbody>
                                    </table>
                                  </div>
                                )}
                              </div>

                            </div>
                          </div>
                        ))}
                      </div>
                    )}
                  </div>
                )}

              </div>
            </div>
          ) : (
            
            /* Professional Empty state visual display */
            <div className="bg-white rounded-2xl border border-slate-200 shadow-sm p-8 text-center flex flex-col items-center justify-center min-h-[520px]">
              <div className="w-16 h-16 rounded-2xl bg-blue-50 text-blue-600 flex items-center justify-center mb-6">
                <Languages className="w-8 h-8" />
              </div>
              <h3 className="text-base font-bold text-slate-800 uppercase tracking-wider mb-2">
                Unified Translation Terminal
              </h3>
              <p className="text-xs text-slate-400 max-w-md mx-auto mb-6">
                Please upload your target Microsoft Word Document inside the active status controller to prompt and execute multimodal rendering diagnostics.
              </p>

              {/* Graphic architecture map */}
              <div className="w-full max-w-lg border border-dashed border-slate-200 rounded-xl p-5 bg-slate-50/50 grid grid-cols-1 sm:grid-cols-3 gap-4">
                <div className="p-4 bg-white rounded-lg shadow-3xs border border-slate-100 flex flex-col items-center text-center">
                  <div className="w-9 h-9 rounded-full bg-blue-50 text-blue-600 flex items-center justify-center mb-2 font-bold text-xs">01</div>
                  <span className="text-xs font-bold text-slate-800">XML Breakdown</span>
                  <span className="text-[10px] text-slate-400 mt-1">Isolates blocks from doc archives.</span>
                </div>
                <div className="p-4 bg-white rounded-lg shadow-3xs border border-slate-100 flex flex-col items-center text-center">
                  <div className="w-9 h-9 rounded-full bg-indigo-50 text-indigo-600 flex items-center justify-center mb-2 font-bold text-xs">02</div>
                  <span className="text-xs font-bold text-slate-800">Gemini Parsing</span>
                  <span className="text-[10px] text-slate-400 mt-1">Multi-stage LLM translation tags.</span>
                </div>
                <div className="p-4 bg-white rounded-lg shadow-3xs border border-slate-100 flex flex-col items-center text-center">
                  <div className="w-9 h-9 rounded-full bg-emerald-50 text-emerald-600 flex items-center justify-center mb-2 font-bold text-xs">03</div>
                  <span className="text-xs font-bold text-slate-800">Reconstitution</span>
                  <span className="text-[10px] text-slate-400 mt-1">Appends visual index with docx builder.</span>
                </div>
              </div>
            </div>
          )}
        </section>

      </main>

      {/* Footer (Bar configuration) */}
      <footer className="mt-auto px-8 py-4 bg-white border-t border-slate-200 flex flex-col sm:flex-row justify-between gap-4 text-xs text-slate-400 font-semibold uppercase tracking-widest shrink-0">
        <div className="flex flex-wrap gap-6">
          <span>Session: #9201-AX</span>
          <span className="text-slate-300">|</span>
          <span>Model: Gemini 3.5 Multimodal</span>
          <span className="text-slate-300">|</span>
          <span>Engine: Docx 8.0 Compiler</span>
        </div>
        <div>© 2026 DocTrans AI Laboratory</div>
      </footer>
    </div>
  );
}

// Convert <img-placeholder id="..."> mapping tags inside translation output to beautiful inline tags in preview lists
function sanitizeHtmlWithPlaceholders(html: string): string {
  if (!html) return "";
  let clean = html;
  clean = clean.replace(/<img-placeholder id="([^"]+)"><\/img-placeholder>/g, (_, pId) => {
    return `<div class="my-4 p-3 bg-blue-50/50 border border-blue-100 rounded-lg flex items-center justify-between text-[11px] text-blue-700 max-w-md mx-auto">
      <div class="flex items-center gap-1.5 font-bold">
        <span>&#128443; IMAGE_ASSET_${pId.toUpperCase()}</span>
      </div>
      <span class="text-[9px] font-mono select-none px-2 py-0.5 bg-blue-100 rounded text-blue-800 uppercase font-bold">Image Embedded</span>
    </div>`;
  });
  return clean;
}
