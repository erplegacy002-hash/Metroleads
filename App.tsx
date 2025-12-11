import React, { useState } from 'react';
import { Loader2, Download, AlertCircle, FileText } from 'lucide-react';
import FileUpload from './components/FileUpload';
import ImageGallery from './components/ImageGallery';
import { ProcessResponse } from './types';
import { processFile } from './utils/clientProcessor';

const App: React.FC = () => {
  const [file, setFile] = useState<File | null>(null);
  const [isLoading, setIsLoading] = useState(false);
  const [result, setResult] = useState<ProcessResponse | null>(null);
  const [error, setError] = useState<string | null>(null);

  const handleProcess = async () => {
    if (!file) return;

    setIsLoading(true);
    setError(null);
    setResult(null);

    try {
      // NOTE: Using pure client-side processing since terminal access is unavailable.
      const data = await processFile(file);
      setResult(data);
    } catch (err: any) {
      console.error(err);
      setError(err.message || 'An unexpected error occurred during processing.');
    } finally {
      setIsLoading(false);
    }
  };

  return (
    <div className="min-h-screen bg-slate-50 pb-20">
      {/* Branded Header */}
      <header className="bg-white border-b border-amber-200/50 sticky top-0 z-10 shadow-sm">
        <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 h-28 flex items-center justify-center">
          <div className="flex items-center space-x-6 sm:space-x-10">
            {/* Metro Logo */}
            <img 
              src="https://d3uv32fm2waqiz.cloudfront.net/b6d7f27/img/metro-logo.png" 
              alt="Metro Group" 
              className="h-10 sm:h-14 w-auto object-contain"
            />
            
            {/* X Separator */}
            <span className="text-2xl sm:text-4xl font-cinzel-dec text-slate-300 font-bold">X</span>
            
            {/* Legacy Logo */}
            <img 
              src="https://github.com/erplegacy002-hash/testbalkemal/blob/main/LOGO.png?raw=true" 
              alt="Legacy Lifespaces" 
              className="h-10 sm:h-14 w-auto object-contain"
            />
          </div>
        </div>
      </header>

      <main className="max-w-6xl mx-auto px-4 sm:px-6 lg:px-8 py-10">
        
        <div className="text-center mb-12">
          <h2 className="text-3xl sm:text-5xl font-cinzel-dec font-bold text-slate-900 tracking-tight mb-4">
            Daily Report Processor
          </h2>
          <p className="text-lg text-slate-600 max-w-2xl mx-auto font-cinzel">
            Automated formatting for Project Performance Reports (Browser Mode)
          </p>
        </div>

        {/* Upload Section */}
        <FileUpload 
          onFileSelect={setFile} 
          selectedFile={file} 
          disabled={isLoading} 
        />

        {/* Action Area */}
        <div className="flex flex-col items-center justify-center mb-16 space-y-4">
          {error && (
            <div className="flex items-center space-x-2 text-red-600 bg-red-50 px-4 py-2 rounded-lg border border-red-200 shadow-sm">
              <AlertCircle className="w-5 h-5" />
              <span>{error}</span>
            </div>
          )}

          <button
            onClick={handleProcess}
            disabled={!file || isLoading}
            className={`
              flex items-center space-x-3 px-10 py-4 rounded-none font-cinzel font-bold text-lg tracking-wider shadow-md transition-all border
              ${!file || isLoading 
                ? 'bg-slate-300 text-slate-500 border-slate-300 cursor-not-allowed' 
                : 'bg-[#1a1a1a] text-[#d4af37] border-[#d4af37] hover:bg-black hover:shadow-xl hover:scale-105'}
            `}
          >
            {isLoading ? (
              <>
                <Loader2 className="w-5 h-5 animate-spin text-[#d4af37]" />
                <span>Processing...</span>
              </>
            ) : (
              <>
                <FileText className="w-5 h-5" />
                <span>Generate Reports</span>
              </>
            )}
          </button>
        </div>

        {/* Results Section */}
        {result && (
          <div className="space-y-8 animate-in fade-in slide-in-from-bottom-4 duration-700">
            <div className="flex items-center justify-between border-b border-amber-200 pb-4">
              <h3 className="text-2xl font-cinzel-dec font-bold text-slate-800">
                Generated Tables ({result.images.length})
              </h3>
              <a
                href={result.zip_url}
                download="project_reports.zip"
                className="flex items-center space-x-2 bg-[#d4af37] text-black px-6 py-2.5 rounded-sm hover:bg-[#c5a028] transition-colors shadow-sm font-cinzel font-bold text-sm uppercase tracking-wide"
              >
                <Download className="w-4 h-4" />
                <span>Download ZIP</span>
              </a>
            </div>

            <ImageGallery images={result.images} />
          </div>
        )}
      </main>
    </div>
  );
};

export default App;