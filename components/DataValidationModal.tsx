
import React, { useState } from 'react';
import { X, Sparkles, ShieldAlert, CheckCircle, Trash2, PlusCircle } from 'lucide-react';
import { ValidationRule, ValidationType, ValidationOperator } from '../types';
import { generateValidationRule } from '../geminiService';

interface DataValidationModalProps {
  activeCell: string | null;
  currentRule: ValidationRule | null;
  onSet: (rule: ValidationRule | null) => void;
  onClose: () => void;
}

const DataValidationModal: React.FC<DataValidationModalProps> = ({ activeCell, currentRule, onSet, onClose }) => {
  const [prompt, setPrompt] = useState('');
  const [isGenerating, setIsGenerating] = useState(false);
  const [manualType, setManualType] = useState<ValidationType>('list');
  const [listItems, setListItems] = useState<string>('');
  const [errorMessage, setErrorMessage] = useState<string>('Invalid input.');

  const handleAiSuggest = async () => {
    if (!prompt.trim()) return;
    setIsGenerating(true);
    const rule = await generateValidationRule(prompt);
    if (rule) {
      onSet({
        type: rule.type as ValidationType,
        operator: rule.operator as ValidationOperator,
        value1: rule.value1,
        listItems: rule.listItems,
        allowBlank: true,
        errorMessage: rule.errorMessage
      });
      setPrompt('');
    }
    setIsGenerating(false);
  };

  const handleManualSet = () => {
    if (manualType === 'list') {
      const items = listItems.split(',').map(s => s.trim()).filter(Boolean);
      onSet({
        type: 'list',
        listItems: items,
        allowBlank: true,
        errorMessage: errorMessage
      });
    } else {
      onSet({
        type: manualType,
        allowBlank: true,
        errorMessage: errorMessage
      });
    }
  };

  const clearRule = () => {
    onSet(null);
  };

  if (!activeCell) return null;

  return (
    <div className="fixed inset-0 bg-slate-900/40 backdrop-blur-sm z-[100] flex items-center justify-center p-4">
      <div className="bg-white rounded-3xl w-full max-w-2xl shadow-2xl border border-slate-200 overflow-hidden animate-in zoom-in-95 duration-200 flex flex-col max-h-[90vh]">
        <div className="p-6 border-b border-slate-100 flex items-center justify-between shrink-0">
          <div>
            <h2 className="text-lg font-black text-slate-800 flex items-center gap-2 uppercase tracking-tight">
              <ShieldAlert className="w-5 h-5 text-indigo-600" />
              Data Validation
            </h2>
            <p className="text-xs font-medium text-slate-500 mt-0.5 uppercase tracking-tighter">Integrity Check for {activeCell}</p>
          </div>
          <button onClick={onClose} className="p-2 hover:bg-slate-100 rounded-xl transition-colors">
            <X className="w-5 h-5 text-slate-400" />
          </button>
        </div>

        <div className="p-6 space-y-8 overflow-y-auto flex-1">
          <div className="space-y-3">
            <label className="text-[10px] font-black uppercase tracking-widest text-slate-400 block px-1">Configure with AI</label>
            <div className="flex gap-2">
              <input 
                type="text" 
                className="flex-1 bg-slate-50 border border-slate-200 rounded-xl px-4 py-2.5 text-sm focus:ring-2 focus:ring-indigo-500 transition-all outline-none"
                placeholder="e.g. 'Create a dropdown with Yes, No, Maybe'"
                value={prompt}
                onChange={(e) => setPrompt(e.target.value)}
              />
              <button 
                onClick={handleAiSuggest}
                disabled={isGenerating || !prompt}
                className="bg-indigo-600 text-white px-4 py-2 rounded-xl text-xs font-bold hover:bg-indigo-700 transition-all shadow-lg shadow-indigo-200 disabled:opacity-50"
              >
                {isGenerating ? "..." : <Sparkles className="w-4 h-4" />}
              </button>
            </div>
          </div>

          <div className="space-y-4">
            <label className="text-[10px] font-black uppercase tracking-widest text-slate-400 block px-1">Manual Configuration</label>
            <div className="p-5 bg-slate-50 border border-slate-200 rounded-2xl space-y-4">
              <div className="grid grid-cols-2 gap-4">
                <div>
                  <label className="text-[10px] font-bold text-slate-500 uppercase mb-1 block">Criteria</label>
                  <select 
                    className="w-full bg-white border border-slate-200 rounded-lg px-3 py-2 text-sm outline-none focus:ring-2 focus:ring-indigo-500"
                    value={manualType}
                    onChange={(e) => setManualType(e.target.value as ValidationType)}
                  >
                    <option value="list">Dropdown (List of items)</option>
                    <option value="wholeNumber">Whole Number</option>
                    <option value="decimal">Decimal</option>
                    <option value="date">Date</option>
                    <option value="textLength">Text Length</option>
                  </select>
                </div>
                {manualType === 'list' && (
                  <div>
                    <label className="text-[10px] font-bold text-slate-500 uppercase mb-1 block">Items (comma separated)</label>
                    <input 
                      type="text" 
                      className="w-full bg-white border border-slate-200 rounded-lg px-3 py-2 text-sm outline-none focus:ring-2 focus:ring-indigo-500"
                      placeholder="Apple, Banana, Orange"
                      value={listItems}
                      onChange={(e) => setListItems(e.target.value)}
                    />
                  </div>
                )}
              </div>
              <div>
                <label className="text-[10px] font-bold text-slate-500 uppercase mb-1 block">Error Message</label>
                <input 
                  type="text" 
                  className="w-full bg-white border border-slate-200 rounded-lg px-3 py-2 text-sm outline-none focus:ring-2 focus:ring-indigo-500"
                  placeholder="Invalid input."
                  value={errorMessage}
                  onChange={(e) => setErrorMessage(e.target.value)}
                />
              </div>
              <button 
                onClick={handleManualSet}
                className="w-full bg-slate-800 text-white px-4 py-2 rounded-xl text-xs font-bold hover:bg-slate-900 transition-all"
              >
                Apply Manual Rule
              </button>
            </div>
          </div>

          <div className="space-y-4">
            <label className="text-[10px] font-black uppercase tracking-widest text-slate-400 block px-1">Applied Integrity Rule</label>
            {currentRule ? (
              <div className="p-5 bg-indigo-50/50 border border-indigo-100 rounded-2xl relative group">
                <div className="flex items-start gap-4">
                  <div className="p-3 bg-indigo-600 text-white rounded-xl shadow-lg shadow-indigo-100">
                    <CheckCircle className="w-5 h-5" />
                  </div>
                  <div className="flex-1">
                    <h4 className="text-xs font-black text-slate-800 uppercase tracking-tight">{currentRule.type} Rule</h4>
                    <p className="text-[11px] text-slate-500 mt-1 font-medium italic">"{currentRule.errorMessage}"</p>
                    <div className="mt-3 flex items-center gap-2 flex-wrap">
                      <span className="text-[9px] font-bold px-2 py-0.5 bg-indigo-100 text-indigo-700 rounded-full uppercase">{currentRule.operator || 'Custom'}</span>
                      {currentRule.value1 && <span className="text-[9px] font-bold px-2 py-0.5 bg-white text-slate-600 border border-slate-200 rounded-full">{currentRule.value1}</span>}
                      {currentRule.listItems && currentRule.listItems.map((item, i) => (
                        <span key={i} className="text-[9px] font-bold px-2 py-0.5 bg-white text-slate-600 border border-slate-200 rounded-full">{item}</span>
                      ))}
                    </div>
                  </div>
                </div>
                <button 
                  onClick={clearRule}
                  className="absolute top-4 right-4 p-2 text-slate-300 hover:text-rose-500 hover:bg-rose-50 rounded-lg transition-all opacity-0 group-hover:opacity-100"
                >
                  <Trash2 className="w-4 h-4" />
                </button>
              </div>
            ) : (
              <div className="py-12 text-center bg-slate-50 border border-dashed border-slate-200 rounded-2xl">
                <ShieldAlert className="w-10 h-10 text-slate-200 mx-auto mb-3" />
                <p className="text-xs font-medium text-slate-400 px-10">
                  No active validation. Use the AI bar above to define entry constraints.
                </p>
              </div>
            )}
          </div>
        </div>

        <div className="p-6 bg-slate-50 border-t border-slate-100 flex justify-end shrink-0">
          <button onClick={onClose} className="px-8 py-3 bg-slate-900 text-white rounded-2xl text-xs font-black hover:bg-indigo-600 transition-all shadow-xl shadow-indigo-100 uppercase tracking-widest">
            Lock Settings
          </button>
        </div>
      </div>
    </div>
  );
};

export default DataValidationModal;
