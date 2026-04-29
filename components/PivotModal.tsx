import React, { useState, useEffect } from 'react';
import { X, TableProperties, Loader2 } from 'lucide-react';
import { SheetState } from '../types';
import Markdown from 'react-markdown';

interface PivotModalProps {
  sheet: SheetState;
  onClose: () => void;
}

const PivotModal: React.FC<PivotModalProps> = ({ sheet, onClose }) => {
  const [pivotData, setPivotData] = useState<string | null>(null);
  const [isLoading, setIsLoading] = useState(true);

  const parseCellId = (id: string) => {
    const colMatch = id.match(/[A-Z]+/);
    const rowMatch = id.match(/[0-9]+/);
    if (!colMatch || !rowMatch) return { r: 0, c: 0 };
    const col = colMatch[0];
    const row = parseInt(rowMatch[0]);
    let colIndex = 0;
    for (let i = 0; i < col.length; i++) {
      colIndex = colIndex * 26 + (col.charCodeAt(i) - 64);
    }
    return { r: row, c: colIndex };
  };

  const getCellIdFromCoords = (r: number, c: number) => {
    let col = "";
    let tempC = c;
    while (tempC > 0) {
      let mod = (tempC - 1) % 26;
      col = String.fromCharCode(65 + mod) + col;
      tempC = Math.floor((tempC - mod) / 26);
    }
    return `${col}${r}`;
  };

  useEffect(() => {
    const generatePivot = async () => {
      if (!sheet.selectionStart || !sheet.selectionEnd) {
        setPivotData("Please select a range of data first.");
        setIsLoading(false);
        return;
      }

      const start = parseCellId(sheet.selectionStart);
      const end = parseCellId(sheet.selectionEnd);

      const r1 = Math.min(start.r, end.r);
      const r2 = Math.max(start.r, end.r);
      const c1 = Math.min(start.c, end.c);
      const c2 = Math.max(start.c, end.c);

      const contextData = [];
      for (let r = r1; r <= r2; r++) {
        const row = [];
        for (let c = c1; c <= c2; c++) {
          const cellId = getCellIdFromCoords(r, c);
          row.push(sheet.cells[cellId]?.displayValue || '');
        }
        contextData.push(row.join(', '));
      }

      const prompt = `Analyze the following spreadsheet data and create a summarized Pivot Table view. 
Identify the most likely categorical columns to group by, and the numerical columns to aggregate (sum/average).
Present the result as a clean Markdown table.

Data:
${contextData.join('\n')}`;

      try {
        const { GoogleGenAI } = await import('@google/genai');
        const ai = new GoogleGenAI({ apiKey: import.meta.env.VITE_GEMINI_API_KEY || 'dummy' });
        const response = await ai.models.generateContent({
          model: 'gemini-2.5-flash',
          contents: prompt,
        });
        
        setPivotData(response.text || "Could not generate pivot table.");
      } catch (e) {
        setPivotData("Error generating pivot table.");
      }
      setIsLoading(false);
    };

    generatePivot();
  }, [sheet]);

  return (
    <div className="fixed inset-0 bg-black/50 flex items-center justify-center z-50 backdrop-blur-sm p-4">
      <div className="bg-white rounded-2xl shadow-2xl w-full max-w-3xl max-h-[80vh] flex flex-col overflow-hidden animate-in fade-in zoom-in-95 duration-200">
        <div className="bg-slate-900 p-4 flex items-center justify-between shrink-0">
          <div className="flex items-center gap-2 text-white">
            <TableProperties className="w-5 h-5" />
            <h3 className="font-bold text-sm uppercase tracking-wide">AI Pivot Table</h3>
          </div>
          <button onClick={onClose} className="text-slate-400 hover:text-white transition-colors">
            <X className="w-5 h-5" />
          </button>
        </div>

        <div className="flex-1 p-6 overflow-y-auto bg-slate-50">
          {isLoading ? (
            <div className="flex flex-col items-center justify-center h-full text-slate-500 gap-4">
              <Loader2 className="w-8 h-8 animate-spin text-indigo-500" />
              <p className="font-medium animate-pulse">Analyzing data and generating pivot table...</p>
            </div>
          ) : (
            <div className="prose prose-sm max-w-none prose-slate prose-tables:w-full prose-tables:border-collapse prose-th:bg-slate-200 prose-th:p-2 prose-th:border prose-th:border-slate-300 prose-td:p-2 prose-td:border prose-td:border-slate-300 bg-white p-6 rounded-xl shadow-sm border border-slate-200">
              <Markdown>{pivotData || ''}</Markdown>
            </div>
          )}
        </div>
      </div>
    </div>
  );
};

export default PivotModal;
