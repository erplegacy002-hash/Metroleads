import React, { useState } from 'react';
import { Loader2, Download, AlertCircle, FileText, MapPin, CalendarRange, CalendarDays, Calendar } from 'lucide-react';
import FileUpload from './components/FileUpload';
import ImageGallery from './components/ImageGallery';
import { ProcessResponse } from './types';
import { processFile } from './utils/clientProcessor';
import { processMonthlyFile } from './utils/monthlyProcessor';
import { processDailySiteVisitFile } from './utils/dailySiteVisitProcessor';
import { processWeeklySiteVisitFile } from './utils/weeklySiteVisitProcessor';

const App: React.FC = () => {
  const [activeTab, setActiveTab] = useState('Daily Report Processor');
  const [file, setFile] = useState<File | null>(null);
  const [isLoading, setIsLoading] = useState(false);
  const [result, setResult] = useState<ProcessResponse | null>(null);
  const [error, setError] = useState<string | null>(null);
  
  // Date states
  const [startDate, setStartDate] = useState('');
  const [endDate, setEndDate] = useState('');

  const handleProcess = async () => {
    if (!file) return;

    setIsLoading(true);
    setError(null);
    setResult(null);

    try {
      let data: ProcessResponse;
      
      if (activeTab === 'Monthly Site Visit Report') {
        data = await processMonthlyFile(file, startDate, endDate);
      } else if (activeTab === 'Weekly Site Visit Report') {
        data = await processWeeklySiteVisitFile(file, startDate, endDate);
      } else if (activeTab === 'Daily Site Visit Report') {
        data = await processDailySiteVisitFile(file, startDate, endDate);
      } else {
        // Default to Daily Report Processor
        data = await processFile(file);
      }
      
      setResult(data);
    } catch (err: any) {
      console.error(err);
      setError(err.message || 'An unexpected error occurred during processing.');
    } finally {
      setIsLoading(false);
    }
  };

  const handleTabChange = (tabId: string) => {
    setActiveTab(tabId);
    setFile(null);
    setResult(null);
    setError(null);

    const yesterday = new Date();
    yesterday.setDate(yesterday.getDate() - 1);
    
    const formatDate = (date: Date) => {
      const year = date.getFullYear();
      const month = String(date.getMonth() + 1).padStart(2, '0');
      const day = String(date.getDate()).padStart(2, '0');
      return `${year}-${month}-${day}`;
    };

    if (tabId === 'Daily Site Visit Report') {
      const formattedDate = formatDate(yesterday);
      setStartDate(formattedDate);
      setEndDate(formattedDate);
    } else if (tabId === 'Weekly Site Visit Report') {
      const formattedEnd = formatDate(yesterday);
      
      const lastWeek = new Date(yesterday);
      lastWeek.setDate(yesterday.getDate() - 6); // 7 days inclusive range
      const formattedStart = formatDate(lastWeek);

      setStartDate(formattedStart);
      setEndDate(formattedEnd);
    } else {
      setStartDate('');
      setEndDate('');
    }
  };

  const tabs = [
    { id: 'Daily Report Processor', label: 'Daily Report Processor', icon: FileText },
    { id: 'Daily Site Visit Report', label: 'Daily Site Visit Report', icon: MapPin },
    { id: 'Weekly Site Visit Report', label: 'Weekly Site Visit Report', icon: CalendarDays },
    { id: 'Monthly Site Visit Report', label: 'Monthly Site Visit Report', icon: CalendarRange },
  ];

  const isProcessorTab = activeTab === 'Daily Report Processor' || activeTab === 'Monthly Site Visit Report' || activeTab === 'Daily Site Visit Report' || activeTab === 'Weekly Site Visit Report';
  const showDateInputs = activeTab !== 'Daily Report Processor';

  return (
    <div className="min-h-screen bg-slate-50 pb-20 font-cinzel">
      {/* Branded Header - Logos Centered, Badge Top Right */}
      <header className="bg-white border-b border-amber-200/50 sticky top-0 z-30 shadow-sm">
        <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 h-24 sm:h-32 flex items-center justify-center relative">
          
          {/* Main Logo Cluster (Centered) */}
          <div className="flex items-center space-x-4 sm:space-x-12 scale-90 sm:scale-100">
            {/* Metro Logo */}
            <img 
              src="https://d3uv32fm2waqiz.cloudfront.net/b6d7f27/img/metro-logo.png" 
              alt="Metro Group" 
              className="h-8 sm:h-14 w-auto object-contain"
            />
            
            {/* X Separator */}
            <span className="text-xl sm:text-4xl font-cinzel-dec text-slate-300 font-bold">X</span>
            
            {/* Legacy Logo - Using provided GitHub raw format */}
            <img 
              src="https://github.com/erplegacy002-hash/testbalkemal/blob/main/LOGO.png?raw=true" 
              alt="Legacy Lifespaces" 
              className="h-8 sm:h-14 w-auto object-contain"
            />
          </div>

          {/* GPTW Badge (Top Right) */}
          <div className="absolute top-2 right-2 sm:top-4 sm:right-4">
            <img 
              src="https://www.greatplacetowork.in/great/api/assets/uploads/13713/logo/batch.png" 
              alt="Great Place to Work" 
              className="h-10 sm:h-20 w-auto object-contain"
            />
          </div>
        </div>
      </header>

      {/* Tabs Navigation */}
      <div className="bg-white border-b border-slate-200 sticky top-24 sm:top-32 z-20 shadow-sm overflow-x-auto scrollbar-hide">
        <div className="max-w-6xl mx-auto px-4 sm:px-6 lg:px-8">
          <nav className="-mb-px flex space-x-8" aria-label="Tabs">
            {tabs.map((tab) => {
              const Icon = tab.icon;
              const isActive = activeTab === tab.id;
              return (
                <button
                  key={tab.id}
                  onClick={() => handleTabChange(tab.id)}
                  className={`
                    whitespace-nowrap py-4 px-1 border-b-2 font-medium text-sm flex items-center space-x-2 transition-colors outline-none
                    ${isActive 
                      ? 'border-[#d4af37] text-[#1a1a1a]' 
                      : 'border-transparent text-slate-500 hover:text-slate-700 hover:border-slate-300'}
                  `}
                >
                  <Icon className={`w-4 h-4 ${isActive ? 'text-[#d4af37]' : ''}`} />
                  <span className="font-inter uppercase tracking-wide text-xs sm:text-sm font-bold">{tab.label}</span>
                </button>
              );
            })}
          </nav>
        </div>
      </div>

      <main className="max-w-6xl mx-auto px-4 sm:px-6 lg:px-8 py-10">
        
        {isProcessorTab ? (
          <div className="animate-in fade-in duration-500">
            <div className="text-center mb-12">
              <h2 className="text-3xl sm:text-5xl font-cinzel-dec font-bold text-slate-900 tracking-tight mb-4">
                {activeTab}
              </h2>
              <p className="text-lg text-slate-600 max-w-2xl mx-auto font-inter">
                {activeTab === 'Monthly Site Visit Report' 
                  ? 'Automated monthly site visit summaries grouped by project.'
                  : activeTab === 'Daily Site Visit Report'
                    ? 'Automated daily site visit reports grouped by project.'
                    : activeTab === 'Weekly Site Visit Report'
                      ? 'Automated weekly site visit reports grouped by project.'
                      : 'Automated formatting for Project Performance Reports (Browser Mode)'}
              </p>
            </div>

            {/* Date Inputs for Site Visit Reports */}
            {showDateInputs && (
              <div className="flex flex-col sm:flex-row items-center justify-center gap-4 mb-8 max-w-2xl mx-auto">
                <div className="w-full">
                  <label className="block text-sm font-semibold text-slate-700 mb-1 font-inter">Start Date</label>
                  <div className="relative">
                    <div className="absolute inset-y-0 left-0 pl-3 flex items-center pointer-events-none">
                      <Calendar className="h-4 w-4 text-slate-400" />
                    </div>
                    <input
                      type="date"
                      value={startDate}
                      onChange={(e) => setStartDate(e.target.value)}
                      className="block w-full pl-10 pr-3 py-2.5 border border-slate-300 rounded-lg focus:ring-[#d4af37] focus:border-[#d4af37] text-sm font-sans"
                    />
                  </div>
                </div>
                <div className="w-full">
                  <label className="block text-sm font-semibold text-slate-700 mb-1 font-inter">End Date</label>
                  <div className="relative">
                    <div className="absolute inset-y-0 left-0 pl-3 flex items-center pointer-events-none">
                      <Calendar className="h-4 w-4 text-slate-400" />
                    </div>
                    <input
                      type="date"
                      value={endDate}
                      onChange={(e) => setEndDate(e.target.value)}
                      className="block w-full pl-10 pr-3 py-2.5 border border-slate-300 rounded-lg focus:ring-[#d4af37] focus:border-[#d4af37] text-sm font-sans"
                    />
                  </div>
                </div>
              </div>
            )}

            {/* Upload Section */}
            <FileUpload 
              onFileSelect={setFile} 
              selectedFile={file} 
              disabled={isLoading} 
            />

            {/* Action Area */}
            <div className="flex flex-col items-center justify-center mb-16 space-y-4">
              {error && (
                <div className="flex items-center space-x-2 text-red-600 bg-red-50 px-4 py-2 rounded-lg border border-red-200 shadow-sm font-sans">
                  <AlertCircle className="w-5 h-5" />
                  <span>{error}</span>
                </div>
              )}

              <button
                onClick={handleProcess}
                disabled={!file || isLoading}
                className={`
                  flex items-center space-x-3 px-10 py-4 rounded-none font-bold text-lg tracking-wider shadow-md transition-all border
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
                    className="flex items-center space-x-2 bg-[#d4af37] text-black px-6 py-2.5 rounded-sm hover:bg-[#c5a028] transition-colors shadow-sm font-bold text-sm uppercase tracking-wide"
                  >
                    <Download className="w-4 h-4" />
                    <span>Download ZIP</span>
                  </a>
                </div>

                <ImageGallery images={result.images} />
              </div>
            )}
          </div>
        ) : (
           <div className="flex flex-col items-center justify-center py-20 text-slate-400 bg-white rounded-xl border border-dashed border-slate-300 animate-in fade-in duration-500">
              <div className="p-6 bg-slate-50 rounded-full mb-6">
                 {activeTab.includes('Daily Site') && <MapPin className="w-12 h-12 text-slate-300" />}
                 {activeTab.includes('Weekly') && <CalendarDays className="w-12 h-12 text-slate-300" />}
              </div>
              <h3 className="text-2xl font-cinzel font-bold text-slate-600 mb-2">{activeTab}</h3>
              <p className="font-inter text-slate-500">This module is under development.</p>
           </div>
        )}

      </main>
    </div>
  );
};

export default App;