import React, { useCallback, useState } from 'react';
import { Upload, FileSpreadsheet, X } from 'lucide-react';

interface FileUploadProps {
  onFileSelect: (file: File | null) => void;
  selectedFile: File | null;
  disabled: boolean;
}

const FileUpload: React.FC<FileUploadProps> = ({ onFileSelect, selectedFile, disabled }) => {
  const [isDragging, setIsDragging] = useState(false);

  const handleDragOver = useCallback((e: React.DragEvent) => {
    e.preventDefault();
    if (!disabled) setIsDragging(true);
  }, [disabled]);

  const handleDragLeave = useCallback((e: React.DragEvent) => {
    e.preventDefault();
    if (!disabled) setIsDragging(false);
  }, [disabled]);

  const handleDrop = useCallback((e: React.DragEvent) => {
    e.preventDefault();
    setIsDragging(false);
    if (disabled) return;

    if (e.dataTransfer.files && e.dataTransfer.files.length > 0) {
      const file = e.dataTransfer.files[0];
      // regex to match xlsx, xls, csv, xlsb, ods, xlsm
      if (file.name.match(/\.(xlsx|xls|csv|xlsb|ods|xlsm)$/i)) {
        onFileSelect(file);
      } else {
        alert("Please upload a valid spreadsheet file (.xlsx, .xls, .csv, .xlsb, .ods, .xlsm)");
      }
    }
  }, [onFileSelect, disabled]);

  const handleFileInput = (e: React.ChangeEvent<HTMLInputElement>) => {
    if (e.target.files && e.target.files.length > 0) {
      onFileSelect(e.target.files[0]);
    }
  };

  const removeFile = () => {
    onFileSelect(null);
  };

  return (
    <div className="w-full max-w-2xl mx-auto mb-8">
      {!selectedFile ? (
        <label
          onDragOver={handleDragOver}
          onDragLeave={handleDragLeave}
          onDrop={handleDrop}
          className={`flex flex-col items-center justify-center w-full h-64 border-2 border-dashed rounded-lg cursor-pointer transition-colors duration-200 
            ${isDragging ? 'border-blue-500 bg-blue-50' : 'border-slate-300 bg-white hover:bg-slate-50'}
            ${disabled ? 'opacity-50 cursor-not-allowed' : ''}
          `}
        >
          <div className="flex flex-col items-center justify-center pt-5 pb-6 text-center">
            <Upload className={`w-12 h-12 mb-4 ${isDragging ? 'text-blue-500' : 'text-slate-400'}`} />
            <p className="mb-2 text-sm text-slate-600">
              <span className="font-semibold">Click to upload</span> or drag and drop
            </p>
            <p className="text-xs text-slate-500">Supports .xlsx, .xls, .csv, .xlsb, .ods, .xlsm</p>
          </div>
          <input 
            type="file" 
            className="hidden" 
            accept=".xlsx,.xls,.csv,.xlsb,.ods,.xlsm,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,application/vnd.ms-excel,text/csv" 
            onChange={handleFileInput} 
            disabled={disabled}
          />
        </label>
      ) : (
        <div className="bg-white border border-slate-200 rounded-lg p-6 flex items-center justify-between shadow-sm">
          <div className="flex items-center space-x-4">
            <div className="p-3 bg-green-100 rounded-full">
              <FileSpreadsheet className="w-6 h-6 text-green-600" />
            </div>
            <div>
              <p className="text-sm font-medium text-slate-900">{selectedFile.name}</p>
              <p className="text-xs text-slate-500">{(selectedFile.size / 1024).toFixed(2)} KB</p>
            </div>
          </div>
          {!disabled && (
            <button 
              onClick={removeFile}
              className="p-2 hover:bg-slate-100 rounded-full text-slate-500 hover:text-red-500 transition-colors"
            >
              <X className="w-5 h-5" />
            </button>
          )}
        </div>
      )}
    </div>
  );
};

export default FileUpload;