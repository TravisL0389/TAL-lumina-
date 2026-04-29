
import React, { useState, useEffect, useRef } from 'react';
import { CellId, CellData, SheetState } from '../types';
import { Sparkles, FunctionSquare, AlertTriangle, CheckCircle2 } from 'lucide-react';
import { suggestFormula, checkFormulaErrors, getFormulaPrediction, explainFormula } from '../geminiService';
import { Loader2, HelpCircle } from 'lucide-react';

interface FormulaBarProps {
  activeCell: CellId | null;
  cellData: CellData | null;
  updateCell: (id: CellId, data: Partial<CellData>) => void;
  sheet: SheetState;
}

const FormulaBar: React.FC<FormulaBarProps> = ({ activeCell, cellData, updateCell, sheet }) => {
  const [inputValue, setInputValue] = useState(cellData?.formula || cellData?.value || '');
  const [isSuggesting, setIsSuggesting] = useState(false);
  const [isAiMode, setIsAiMode] = useState(false);
  const [suggestion, setSuggestion] = useState<string | null>(null);
  const [prediction, setPrediction] = useState<string | null>(null);
  const [explanation, setExplanation] = useState<string | null>(null);
  const [isExplaining, setIsExplaining] = useState(false);
  const checkTimeout = useRef<number | null>(null);
  const predictionTimeout = useRef<number | null>(null);

  useEffect(() => {
    return () => {
      if (checkTimeout.current) clearTimeout(checkTimeout.current);
      if (predictionTimeout.current) clearTimeout(predictionTimeout.current);
    };
  }, []);

  useEffect(() => {
    setInputValue(cellData?.formula || cellData?.value || '');
    setSuggestion(null);
    setPrediction(null);
    setExplanation(null);
    setIsAiMode(false);
  }, [cellData, activeCell]);

  const handleExplain = async () => {
    if (!inputValue || !inputValue.startsWith('=')) return;
    setIsExplaining(true);
    const context = `Active Cell: ${activeCell}`;
    const result = await explainFormula(inputValue, context);
    setExplanation(result);
    setIsExplaining(false);
  };

  const handleBlur = () => {
    if (activeCell) {
      if (inputValue.startsWith('=') || inputValue.startsWith('=AI(')) {
        updateCell(activeCell, { formula: inputValue, displayValue: "Evaluating..." });
      } else {
        updateCell(activeCell, { value: inputValue, displayValue: inputValue, formula: '' });
      }
    }
    setPrediction(null);
  };

  const handleChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    const val = e.target.value;
    setInputValue(val);
    setPrediction(null);
    
    // Auto-check for formula errors with debounce
    if (val.startsWith('=')) {
      if (checkTimeout.current) clearTimeout(checkTimeout.current);
      checkTimeout.current = window.setTimeout(async () => {
        const errorSuggestion = await checkFormulaErrors(val, "Sheet data context for current cell.");
        setSuggestion(errorSuggestion);
      }, 800);
    } else {
      setSuggestion(null);
    }

    // Get AI Prediction
    if (val.length > 2) {
      if (predictionTimeout.current) clearTimeout(predictionTimeout.current);
      predictionTimeout.current = window.setTimeout(async () => {
        const context = `Headers in Row 1. Active Cell: ${activeCell}`;
        const pred = await getFormulaPrediction(val, context);
        if (pred && pred !== val && pred.startsWith('=')) {
          setPrediction(pred);
        }
      }, 600);
    }
  };

  const handleSuggest = async () => {
    if (!activeCell || !inputValue) return;
    setIsSuggesting(true);
    const context = `Spreadsheet with headers in row 1. Current cell is ${activeCell}. Input intent: ${inputValue}`;
    const formula = await suggestFormula(inputValue, context);
    if (formula) {
      setInputValue(formula);
      setIsAiMode(false);
      setPrediction(null);
    }
    setIsSuggesting(false);
  };

  const handleKeyDown = (e: React.KeyboardEvent) => {
    if (e.key === 'Enter') {
      if (isAiMode) {
        handleSuggest();
      } else {
        handleBlur();
      }
    }
    if (e.key === 'Tab' && prediction) {
      e.preventDefault();
      setInputValue(prediction);
      setPrediction(null);
    }
  };

  return (
    <div className={`formula-shell relative flex flex-wrap items-center gap-3 rounded-[1rem] border border-[#e6ebf2] bg-white px-3 py-2 transition-colors ${isAiMode ? 'bg-[#f5fbff]' : ''}`}>
      <div className="flex min-h-9 min-w-[4.25rem] items-center justify-center rounded-lg border border-[#e3e8ef] bg-[#fbfcfe] font-mono text-[11px] font-semibold text-[#6f8096] shadow-sm">
        {activeCell || '---'}
      </div>
      
      <div className="hidden h-6 w-px bg-[#e9eef5] sm:block" />

      <button
        onClick={() => setIsAiMode(!isAiMode)}
        className={`min-h-9 min-w-9 rounded-lg transition-all ${isAiMode ? 'bg-[#edf4ff] text-[#1f6fe5] ring-2 ring-[#dbe9ff]' : 'text-slate-400 hover:bg-[#f3f6fb] hover:text-[#1f6fe5]'}`}
        title="AI Formula Generator"
      >
        <Sparkles className="w-4 h-4" />
      </button>

      <div className={`text-slate-400 transition-opacity ${isAiMode ? 'opacity-0 w-0 overflow-hidden' : 'opacity-100'}`}>
        <FunctionSquare className="w-4 h-4" />
      </div>

      <div className="relative flex min-w-0 flex-1 flex-wrap items-center">
        <div className="relative flex min-h-10 w-full items-center">
          <input 
            type="text"
            className={`z-10 h-full min-h-9 w-full rounded-lg border border-[#e3e8ef] bg-[#fbfcfe] px-3 text-sm font-medium outline-none transition-colors ${
              isAiMode 
                ? 'text-[#1f6fe5] placeholder:text-[#7aa4e8]' 
                : 'text-slate-700 placeholder:text-slate-400'
            } ${suggestion ? 'text-orange-500' : ''}`}
            placeholder={isAiMode ? "Describe logic (e.g. 'VLOOKUP Apple', 'Flash Fill names', 'Pivot summary')..." : "Type data, =Formula, or natural language..."}
            value={inputValue}
            onChange={handleChange}
            onKeyDown={handleKeyDown}
            onBlur={isAiMode ? undefined : handleBlur}
            autoFocus={isAiMode}
          />
          {prediction && (
            <div className="absolute inset-0 flex items-center pointer-events-none">
              <span className="invisible whitespace-pre">{inputValue}</span>
              <span className="ml-1 truncate text-sm font-medium text-slate-400 opacity-60">
                 {prediction.startsWith(inputValue) ? prediction.slice(inputValue.length) : `  →  ${prediction}`}
                 <span className="ml-2 rounded border border-[#e3e8ef] bg-white px-1 text-[9px] text-slate-400">TAB</span>
              </span>
            </div>
          )}
        </div>

        {suggestion && (
          <div className="absolute left-0 top-[calc(100%+0.5rem)] z-50 flex items-center gap-2 rounded-lg border border-orange-200 bg-[#fff7ed] px-3 py-1.5 text-xs font-semibold text-orange-600 shadow-xl animate-in fade-in slide-in-from-top-1">
            <AlertTriangle className="w-3.5 h-3.5" />
            AI Correction: <span className="underline cursor-pointer" onClick={() => { setInputValue(suggestion); setSuggestion(null); }}>{suggestion}</span>
          </div>
        )}

        {inputValue && !inputValue.startsWith('=') && !isSuggesting && !isAiMode && (
          <button 
            onClick={handleSuggest}
            className="mt-2 inline-flex min-h-9 items-center gap-1.5 rounded-full border border-[#dbe9ff] bg-[#f4f8ff] px-3 py-2 text-[10px] font-semibold text-[#1f6fe5] transition-all hover:bg-[#eaf2ff] sm:mt-0 sm:ml-2"
          >
            <Sparkles className="w-3 h-3" />
            GENERATE FORMULA
          </button>
        )}

        {inputValue && inputValue.startsWith('=') && !isExplaining && !isAiMode && (
          <button 
            onClick={handleExplain}
            className="mt-2 inline-flex min-h-9 items-center gap-1.5 rounded-full border border-emerald-200 bg-emerald-50 px-3 py-2 text-[10px] font-semibold text-emerald-600 transition-all hover:bg-emerald-100 sm:mt-0 sm:ml-2"
          >
            <HelpCircle className="w-3 h-3" />
            EXPLAIN
          </button>
        )}

        {isExplaining && (
          <div className="ml-2 flex items-center gap-2 text-[10px] font-semibold text-emerald-500">
            <div className="animate-spin h-3 w-3 border-2 border-emerald-300 border-t-transparent rounded-full" />
            EXPLAINING...
          </div>
        )}

        {explanation && (
          <div className="absolute right-0 top-[calc(100%+0.75rem)] z-50 w-full max-w-[22rem] rounded-xl border border-[#dce8f8] bg-white p-4 shadow-2xl animate-in fade-in slide-in-from-top-2">
            <div className="mb-2 flex items-center gap-2 font-semibold text-[#1f6fe5]">
              <Sparkles className="w-4 h-4" />
              Formula Explanation
            </div>
            <p className="text-sm leading-relaxed text-slate-600">{explanation}</p>
            <button 
              onClick={() => setExplanation(null)}
              className="mt-3 w-full rounded-lg bg-[#f3f6fb] py-1.5 text-xs font-semibold text-slate-700 transition-colors hover:bg-[#e8eef7]"
            >
              Close
            </button>
          </div>
        )}

        {isSuggesting && (
          <div className="flex items-center gap-2 text-[10px] font-semibold text-[#1f6fe5]">
            <div className="h-3 w-3 animate-spin rounded-full border-2 border-[#1f6fe5] border-t-transparent" />
            CRAFTING...
          </div>
        )}
      </div>
    </div>
  );
};

export default FormulaBar;
