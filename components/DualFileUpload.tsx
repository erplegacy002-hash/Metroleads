import React, { useCallback, useState } from 'react';
import { Upload, FileSpreadsheet, X, CheckCircle2, AlertCircle } from 'lucide-react';

interface DualFileUploadProps {
  leadsFile: File | null;
  onLeadsFileSelect: (file: File | null) => void;
  visitsFile: File | null;
  onVisitsFileSelect: (file: File | null) => void;
  disabled: boolean;
}

const DualFileUpload: React.FC<DualFileUploadProps> = ({
  leadsFile,
  onLeadsFileSelect,
  visitsFile,
  onVisitsFileSelect,
  disabled
}) => {
  const [isDraggingLeads, setIsDraggingLeads] = useState(false);
  const [isDraggingVisits, setIsDraggingVisits] = useState(false);

  const handleLeadsDrop = useCallback((e: React.DragEvent) => {
    e.preventDefault();
    setIsDraggingLeads(false);
    if (disabled) return;
    if (e.dataTransfer.files && e.dataTransfer.files.length > 0) {
      const file = Array.from(e.dataTransfer.files).find(f => f.name.match(/\.(xlsx|xls|csv|xlsb|ods|xlsm)$/i));
      if (file) onLeadsFileSelect(file);
    }
  }, [disabled, onLeadsFileSelect]);

  const handleVisitsDrop = useCallback((e: React.DragEvent) => {
    e.preventDefault();
    setIsDraggingVisits(false);
    if (disabled) return;
    if (e.dataTransfer.files && e.dataTransfer.files.length > 0) {
      const file = Array.from(e.dataTransfer.files).find(f => f.name.match(/\.(xlsx|xls|csv|xlsb|ods|xlsm)$/i));
      if (file) onVisitsFileSelect(file);
    }
  }, [disabled, onVisitsFileSelect]);

  return (
    <div className="w-full max-w-4xl mx-auto mb-8 space-y-4">
      <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
        {/* 1. Leads Excel Upload Card */}
        <div className="bg-white border border-slate-200 rounded-xl p-4 shadow-sm flex flex-col justify-between">
          <div className="flex items-center justify-between mb-3 border-b border-slate-100 pb-2.5">
            <div className="flex items-center space-x-2">
              <span className="w-5 h-5 rounded-full bg-amber-100 text-amber-800 text-xs font-bold flex items-center justify-center">1</span>
              <h4 className="text-sm font-bold text-slate-800">Leads Excel (Presales)</h4>
            </div>
            <span className="text-[10px] uppercase font-bold tracking-wider px-2 py-0.5 rounded bg-rose-50 text-rose-700 border border-rose-200">
              Mandatory
            </span>
          </div>
          <p className="text-xs text-slate-500 mb-3">
            Contains Telecaller, Agency, Lead Level, Mobile No & Date.
          </p>

          {!leadsFile ? (
            <label
              onDragOver={(e) => { e.preventDefault(); if (!disabled) setIsDraggingLeads(true); }}
              onDragLeave={(e) => { e.preventDefault(); if (!disabled) setIsDraggingLeads(false); }}
              onDrop={handleLeadsDrop}
              className={`flex flex-col items-center justify-center w-full h-36 border-2 border-dashed rounded-lg cursor-pointer transition-all duration-200 
                ${isDraggingLeads ? 'border-amber-500 bg-amber-50/50' : 'border-slate-300 bg-slate-50/60 hover:bg-slate-50 hover:border-slate-400'}
                ${disabled ? 'opacity-50 cursor-not-allowed' : ''}
              `}
            >
              <div className="flex flex-col items-center justify-center text-center p-3">
                <Upload className={`w-7 h-7 mb-2 ${isDraggingLeads ? 'text-amber-600' : 'text-slate-400'}`} />
                <p className="text-xs text-slate-700 font-semibold mb-0.5">
                  Click to upload Leads Excel
                </p>
                <p className="text-[11px] text-slate-400">or drag and drop (.xlsx, .csv)</p>
              </div>
              <input 
                type="file" 
                className="hidden" 
                accept=".xlsx,.xls,.csv,.xlsb,.ods,.xlsm,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,application/vnd.ms-excel,text/csv" 
                onChange={(e) => {
                  if (e.target.files && e.target.files[0]) {
                    onLeadsFileSelect(e.target.files[0]);
                  }
                }} 
                disabled={disabled}
              />
            </label>
          ) : (
            <div className="bg-amber-50/40 border border-amber-200/80 rounded-lg p-3 flex items-center justify-between">
              <div className="flex items-center space-x-3 truncate">
                <div className="p-2 bg-amber-100 rounded-lg shrink-0">
                  <FileSpreadsheet className="w-5 h-5 text-amber-700" />
                </div>
                <div className="truncate">
                  <p className="text-xs font-bold text-slate-900 truncate" title={leadsFile.name}>{leadsFile.name}</p>
                  <p className="text-[10px] text-slate-500">{(leadsFile.size / 1024).toFixed(1)} KB • Ready</p>
                </div>
              </div>
              {!disabled && (
                <button 
                  type="button"
                  onClick={() => onLeadsFileSelect(null)}
                  className="p-1 hover:bg-amber-100 rounded-full text-slate-400 hover:text-red-500 transition-colors shrink-0 ml-2"
                  title="Remove file"
                >
                  <X className="w-4 h-4" />
                </button>
              )}
            </div>
          )}
        </div>

        {/* 2. Visits Excel Upload Card */}
        <div className="bg-white border border-slate-200 rounded-xl p-4 shadow-sm flex flex-col justify-between">
          <div className="flex items-center justify-between mb-3 border-b border-slate-100 pb-2.5">
            <div className="flex items-center space-x-2">
              <span className="w-5 h-5 rounded-full bg-blue-100 text-blue-800 text-xs font-bold flex items-center justify-center">2</span>
              <h4 className="text-sm font-bold text-slate-800">Visits Excel (Site Visits)</h4>
            </div>
            <span className="text-[10px] uppercase font-bold tracking-wider px-2 py-0.5 rounded bg-rose-50 text-rose-700 border border-rose-200">
              Mandatory
            </span>
          </div>
          <p className="text-xs text-slate-500 mb-3">
            Contains Mobile No, Date, Revisit Date & Project Name.
          </p>

          {!visitsFile ? (
            <label
              onDragOver={(e) => { e.preventDefault(); if (!disabled) setIsDraggingVisits(true); }}
              onDragLeave={(e) => { e.preventDefault(); if (!disabled) setIsDraggingVisits(false); }}
              onDrop={handleVisitsDrop}
              className={`flex flex-col items-center justify-center w-full h-36 border-2 border-dashed rounded-lg cursor-pointer transition-all duration-200 
                ${isDraggingVisits ? 'border-blue-500 bg-blue-50/50' : 'border-slate-300 bg-slate-50/60 hover:bg-slate-50 hover:border-slate-400'}
                ${disabled ? 'opacity-50 cursor-not-allowed' : ''}
              `}
            >
              <div className="flex flex-col items-center justify-center text-center p-3">
                <Upload className={`w-7 h-7 mb-2 ${isDraggingVisits ? 'text-blue-600' : 'text-slate-400'}`} />
                <p className="text-xs text-slate-700 font-semibold mb-0.5">
                  Click to upload Visits Excel
                </p>
                <p className="text-[11px] text-slate-400">or drag and drop (.xlsx, .csv)</p>
              </div>
              <input 
                type="file" 
                className="hidden" 
                accept=".xlsx,.xls,.csv,.xlsb,.ods,.xlsm,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,application/vnd.ms-excel,text/csv" 
                onChange={(e) => {
                  if (e.target.files && e.target.files[0]) {
                    onVisitsFileSelect(e.target.files[0]);
                  }
                }} 
                disabled={disabled}
              />
            </label>
          ) : (
            <div className="bg-blue-50/40 border border-blue-200/80 rounded-lg p-3 flex items-center justify-between">
              <div className="flex items-center space-x-3 truncate">
                <div className="p-2 bg-blue-100 rounded-lg shrink-0">
                  <FileSpreadsheet className="w-5 h-5 text-blue-700" />
                </div>
                <div className="truncate">
                  <p className="text-xs font-bold text-slate-900 truncate" title={visitsFile.name}>{visitsFile.name}</p>
                  <p className="text-[10px] text-slate-500">{(visitsFile.size / 1024).toFixed(1)} KB • Ready</p>
                </div>
              </div>
              {!disabled && (
                <button 
                  type="button"
                  onClick={() => onVisitsFileSelect(null)}
                  className="p-1 hover:bg-blue-100 rounded-full text-slate-400 hover:text-red-500 transition-colors shrink-0 ml-2"
                  title="Remove file"
                >
                  <X className="w-4 h-4" />
                </button>
              )}
            </div>
          )}
        </div>
      </div>

      {/* Upload Status Banner */}
      {leadsFile && visitsFile ? (
        <div className="flex items-center space-x-2.5 p-3 rounded-lg bg-emerald-50 border border-emerald-200 text-emerald-800 text-xs font-medium">
          <CheckCircle2 className="w-4 h-4 text-emerald-600 shrink-0" />
          <span>Both Excel files uploaded. Visits Done column will be calculated based on mobile numbers and revisit date rules.</span>
        </div>
      ) : (
        <div className="flex items-center space-x-2.5 p-3 rounded-lg bg-amber-50/80 border border-amber-200 text-amber-800 text-xs">
          <AlertCircle className="w-4 h-4 text-amber-600 shrink-0" />
          <span>
            {!leadsFile && !visitsFile && 'Both Leads Excel and Visits Excel are mandatory. Please upload both files.'}
            {leadsFile && !visitsFile && 'Leads Excel uploaded. Please upload Visits Excel to calculate the Visits Done column.'}
            {!leadsFile && visitsFile && 'Visits Excel uploaded. Please upload Leads Excel to proceed.'}
          </span>
        </div>
      )}
    </div>
  );
};

export default DualFileUpload;
