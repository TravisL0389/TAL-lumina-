import React, { useState, useRef } from 'react';
import { X, Upload, FileSpreadsheet, Database, Globe, Loader2 } from 'lucide-react';
import Papa from 'papaparse';

interface ImportModalProps {
  onImport: (data: string[][]) => void;
  onClose: () => void;
}

const ImportModal: React.FC<ImportModalProps> = ({ onImport, onClose }) => {
  const [isImporting, setIsImporting] = useState(false);
  const fileInputRef = useRef<HTMLInputElement>(null);

  const handleFileUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    setIsImporting(true);
    Papa.parse(file, {
      complete: (results) => {
        setIsImporting(false);
        onImport(results.data as string[][]);
        onClose();
      },
      error: (error) => {
        setIsImporting(false);
        alert('Error parsing CSV: ' + error.message);
      }
    });
  };

  return (
    <div className="fixed inset-0 bg-slate-900/50 backdrop-blur-sm z-50 flex items-center justify-center p-4">
      <div className="bg-white rounded-2xl shadow-2xl w-full max-w-2xl overflow-hidden flex flex-col">
        <div className="p-4 border-b border-slate-100 flex justify-between items-center bg-slate-50">
          <h2 className="font-bold text-slate-800 flex items-center gap-2">
            <Database className="w-5 h-5 text-indigo-600" />
            Power Query / Import Data
          </h2>
          <button onClick={onClose} className="p-1 hover:bg-slate-200 rounded-full text-slate-500 transition-colors">
            <X className="w-5 h-5" />
          </button>
        </div>
        
        <div className="p-6 grid grid-cols-1 md:grid-cols-3 gap-4">
          <div 
            onClick={() => fileInputRef.current?.click()}
            className="border-2 border-dashed border-slate-200 rounded-xl p-6 flex flex-col items-center justify-center text-center cursor-pointer hover:border-indigo-400 hover:bg-indigo-50 transition-all group"
          >
            <input 
              type="file" 
              accept=".csv" 
              className="hidden" 
              ref={fileInputRef}
              onChange={handleFileUpload}
            />
            <div className="w-12 h-12 bg-indigo-100 rounded-full flex items-center justify-center mb-3 group-hover:scale-110 transition-transform">
              {isImporting ? <Loader2 className="w-6 h-6 text-indigo-600 animate-spin" /> : <Upload className="w-6 h-6 text-indigo-600" />}
            </div>
            <h3 className="font-bold text-slate-800 mb-1">Upload CSV</h3>
            <p className="text-xs text-slate-500">Import data from a local file</p>
          </div>

          <div className="border-2 border-slate-100 rounded-xl p-6 flex flex-col items-center justify-center text-center opacity-50 cursor-not-allowed">
            <div className="w-12 h-12 bg-slate-100 rounded-full flex items-center justify-center mb-3">
              <Globe className="w-6 h-6 text-slate-400" />
            </div>
            <h3 className="font-bold text-slate-800 mb-1">From Web</h3>
            <p className="text-xs text-slate-500">Coming soon</p>
          </div>

          <div className="border-2 border-slate-100 rounded-xl p-6 flex flex-col items-center justify-center text-center opacity-50 cursor-not-allowed">
            <div className="w-12 h-12 bg-slate-100 rounded-full flex items-center justify-center mb-3">
              <FileSpreadsheet className="w-6 h-6 text-slate-400" />
            </div>
            <h3 className="font-bold text-slate-800 mb-1">From Database</h3>
            <p className="text-xs text-slate-500">Coming soon</p>
          </div>
        </div>
      </div>
    </div>
  );
};

export default ImportModal;
