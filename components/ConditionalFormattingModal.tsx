
import React, { useState } from 'react';
import { X, Sparkles, Trash2, PlusCircle, Check } from 'lucide-react';
import { ConditionalRule } from '../types';
import { generateFormattingRule } from '../geminiService';

interface ConditionalFormattingModalProps {
  rules: ConditionalRule[];
  onAdd: (rule: ConditionalRule) => void;
  onRemove: (id: string) => void;
  onClose: () => void;
}

const ConditionalFormattingModal: React.FC<ConditionalFormattingModalProps> = ({ rules, onAdd, onRemove, onClose }) => {
  const [prompt, setPrompt] = useState('');
  const [isGenerating, setIsGenerating] = useState(false);
  const [manualType, setManualType] = useState<'greaterThan' | 'lessThan' | 'equals' | 'contains'>('greaterThan');
  const [manualThreshold, setManualThreshold] = useState('');
  const [manualColor, setManualColor] = useState('#ef4444');
  const [manualBg, setManualBg] = useState('#fef2f2');

  const handleGenerate = async () => {
    if (!prompt.trim()) return;
    setIsGenerating(true);
    const result = await generateFormattingRule(prompt);
    if (result) {
      onAdd({
        id: Math.random().toString(36).substr(2, 9),
        description: prompt,
        ...result
      });
      setPrompt('');
    }
    setIsGenerating(false);
  };

  const handleManualAdd = () => {
    onAdd({
      id: Math.random().toString(36).substr(2, 9),
      description: `Manual: ${manualType} ${manualThreshold}`,
      conditionType: manualType,
      threshold: manualThreshold,
      style: { color: manualColor, backgroundColor: manualBg, fontWeight: 'bold' }
    });
    setManualThreshold('');
  };

  return (
    <div className="fixed inset-0 bg-slate-900/40 backdrop-blur-sm z-[100] flex items-center justify-center p-4">
      <div className="bg-white rounded-3xl w-full max-w-2xl shadow-2xl border border-slate-200 overflow-hidden animate-in zoom-in-95 duration-200 flex flex-col max-h-[90vh]">
        <div className="p-6 border-b border-slate-100 flex items-center justify-between shrink-0">
          <div>
            <h2 className="text-lg font-black text-slate-800 flex items-center gap-2">
              <Sparkles className="w-5 h-5 text-indigo-600" />
              Conditional Formatting
            </h2>
            <p className="text-xs font-medium text-slate-500 mt-0.5 uppercase tracking-tighter">AI-Driven Style Policies</p>
          </div>
          <button onClick={onClose} className="p-2 hover:bg-slate-100 rounded-xl transition-colors">
            <X className="w-5 h-5 text-slate-400" />
          </button>
        </div>

        <div className="p-6 space-y-6 overflow-y-auto flex-1">
          <div className="space-y-3">
            <label className="text-[10px] font-black uppercase tracking-widest text-slate-400 block px-1">Create New Rule with AI</label>
            <div className="flex gap-2">
              <input 
                type="text" 
                className="flex-1 bg-slate-50 border border-slate-200 rounded-xl px-4 py-2.5 text-sm focus:ring-2 focus:ring-indigo-500 transition-all outline-none"
                placeholder="e.g. 'Make cells red if value is less than 0'"
                value={prompt}
                onChange={(e) => setPrompt(e.target.value)}
              />
              <button 
                onClick={handleGenerate}
                disabled={isGenerating || !prompt}
                className="bg-indigo-600 text-white px-4 py-2 rounded-xl text-xs font-bold hover:bg-indigo-700 transition-all shadow-lg shadow-indigo-200 disabled:opacity-50"
              >
                {isGenerating ? "..." : <PlusCircle className="w-4 h-4" />}
              </button>
            </div>
          </div>

          <div className="space-y-4">
            <label className="text-[10px] font-black uppercase tracking-widest text-slate-400 block px-1">Manual Configuration</label>
            <div className="p-5 bg-slate-50 border border-slate-200 rounded-2xl space-y-4">
              <div className="grid grid-cols-2 gap-4">
                <div>
                  <label className="text-[10px] font-bold text-slate-500 uppercase mb-1 block">Condition</label>
                  <select 
                    className="w-full bg-white border border-slate-200 rounded-lg px-3 py-2 text-sm outline-none focus:ring-2 focus:ring-indigo-500"
                    value={manualType}
                    onChange={(e) => setManualType(e.target.value as any)}
                  >
                    <option value="greaterThan">Greater Than</option>
                    <option value="lessThan">Less Than</option>
                    <option value="equals">Equals</option>
                    <option value="contains">Contains</option>
                  </select>
                </div>
                <div>
                  <label className="text-[10px] font-bold text-slate-500 uppercase mb-1 block">Value</label>
                  <input 
                    type="text" 
                    className="w-full bg-white border border-slate-200 rounded-lg px-3 py-2 text-sm outline-none focus:ring-2 focus:ring-indigo-500"
                    placeholder="e.g. 100 or 'Pending'"
                    value={manualThreshold}
                    onChange={(e) => setManualThreshold(e.target.value)}
                  />
                </div>
              </div>
              <div className="grid grid-cols-2 gap-4">
                <div>
                  <label className="text-[10px] font-bold text-slate-500 uppercase mb-1 block">Text Color</label>
                  <div className="flex items-center gap-2">
                    <input type="color" value={manualColor} onChange={(e) => setManualColor(e.target.value)} className="w-8 h-8 rounded cursor-pointer" />
                    <span className="text-xs text-slate-500 font-mono">{manualColor}</span>
                  </div>
                </div>
                <div>
                  <label className="text-[10px] font-bold text-slate-500 uppercase mb-1 block">Background Color</label>
                  <div className="flex items-center gap-2">
                    <input type="color" value={manualBg} onChange={(e) => setManualBg(e.target.value)} className="w-8 h-8 rounded cursor-pointer" />
                    <span className="text-xs text-slate-500 font-mono">{manualBg}</span>
                  </div>
                </div>
              </div>
              <button 
                onClick={handleManualAdd}
                disabled={!manualThreshold}
                className="w-full bg-slate-800 text-white px-4 py-2 rounded-xl text-xs font-bold hover:bg-slate-900 transition-all disabled:opacity-50"
              >
                Add Manual Rule
              </button>
            </div>
          </div>

          <div className="space-y-3">
            <label className="text-[10px] font-black uppercase tracking-widest text-slate-400 block px-1">Active Rules ({rules.length})</label>
            <div className="max-h-60 overflow-y-auto space-y-2 pr-2 custom-scrollbar">
              {rules.length === 0 ? (
                <div className="py-8 text-center bg-slate-50 rounded-2xl border border-dashed border-slate-200">
                  <p className="text-xs font-medium text-slate-400">No active rules. Describe a condition above.</p>
                </div>
              ) : (
                rules.map(rule => (
                  <div key={rule.id} className="group p-4 bg-white border border-slate-100 rounded-2xl shadow-sm hover:border-indigo-200 transition-all flex items-center justify-between">
                    <div className="flex items-center gap-3">
                      <div className="w-8 h-8 rounded-lg border border-slate-100 flex items-center justify-center shrink-0" style={rule.style}>
                        <Check className="w-4 h-4" />
                      </div>
                      <div>
                        <p className="text-xs font-bold text-slate-800 uppercase tracking-tighter">{rule.description}</p>
                        <p className="text-[10px] text-slate-400 font-medium">Type: {rule.conditionType} | Threshold: {rule.threshold}</p>
                      </div>
                    </div>
                    <button 
                      onClick={() => onRemove(rule.id)}
                      className="p-2 text-slate-300 hover:text-rose-500 hover:bg-rose-50 rounded-lg transition-all opacity-0 group-hover:opacity-100"
                    >
                      <Trash2 className="w-4 h-4" />
                    </button>
                  </div>
                ))
              )}
            </div>
          </div>
        </div>

        <div className="p-6 bg-slate-50 border-t border-slate-100 flex justify-end shrink-0">
          <button onClick={onClose} className="px-6 py-2 bg-slate-900 text-white rounded-xl text-xs font-black hover:bg-indigo-600 transition-all shadow-xl shadow-indigo-100">
            DONE
          </button>
        </div>
      </div>
    </div>
  );
};

export default ConditionalFormattingModal;
