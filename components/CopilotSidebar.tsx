import React, { useState, useRef, useEffect } from 'react';
import { Send, X, Bot, User, Loader2, Sparkles } from 'lucide-react';
import { copilotChat } from '../geminiService';
import { SheetState } from '../types';

interface CopilotSidebarProps {
  onClose: () => void;
  sheet: SheetState;
  onUpdateRange: (range: string, values: string[][]) => void;
  onApplyFormatting: (range: string, style: React.CSSProperties) => void;
}

export const CopilotSidebar: React.FC<CopilotSidebarProps> = ({ onClose, sheet, onUpdateRange, onApplyFormatting }) => {
  const [messages, setMessages] = useState<{ role: 'user' | 'model', text: string }[]>([
    { role: 'model', text: 'Hi! I am Lumina, your AI assistant. How can I help you with your spreadsheet today?' }
  ]);
  const [input, setInput] = useState('');
  const [isLoading, setIsLoading] = useState(false);
  const messagesEndRef = useRef<HTMLDivElement>(null);

  const scrollToBottom = () => {
    messagesEndRef.current?.scrollIntoView({ behavior: 'smooth' });
  };

  useEffect(() => {
    scrollToBottom();
  }, [messages]);

  const buildCsvContext = () => {
    const maxRows = Math.min(sheet.rowCount, 24);
    const maxCols = Math.min(sheet.columnCount, 10);
    const rows: string[] = [];

    for (let r = 1; r <= maxRows; r += 1) {
      const values: string[] = [];
      for (let c = 1; c <= maxCols; c += 1) {
        let col = '';
        let temp = c;
        while (temp > 0) {
          const mod = (temp - 1) % 26;
          col = String.fromCharCode(65 + mod) + col;
          temp = Math.floor((temp - mod) / 26);
        }
        values.push(sheet.cells[`${col}${r}`]?.displayValue || '');
      }
      rows.push(values.join(','));
    }

    return rows.join('\n');
  };

  const handleSend = async () => {
    if (!input.trim() || isLoading) return;

    const userMessage = input.trim();
    setInput('');
    setMessages(prev => [...prev, { role: 'user', text: userMessage }]);
    setIsLoading(true);

    const csvData = buildCsvContext();

    // Get chat history excluding the initial greeting if it's the only one
    const history = messages.length > 1 ? messages : [];

    const response = await copilotChat(userMessage, csvData, history);

    setMessages(prev => [...prev, { role: 'model', text: response.message }]);

    if (response.actions && response.actions.length > 0) {
      response.actions.forEach((action: any) => {
        if (action.type === 'format' && action.range && action.style) {
          onApplyFormatting(action.range, action.style);
        } else if (action.type === 'update' && action.range && action.values) {
          onUpdateRange(action.range, action.values);
        }
      });
    }

    setIsLoading(false);
  };

  return (
      <div className="z-20 flex h-full w-full flex-col border-l border-[#e6ebf2] bg-white text-slate-700 shadow-[-14px_0_40px_rgba(148,163,184,0.18)]">
      <div className="flex items-center justify-between border-b border-[#eef2f7] bg-[#fbfcfe] p-4">
        <div className="flex items-center gap-2 font-semibold text-[#1f6fe5]">
          <Sparkles size={18} />
          <span>Lumina Copilot</span>
        </div>
        <button onClick={onClose} className="text-slate-400 transition-colors hover:text-slate-700">
          <X size={18} />
        </button>
      </div>

      <div className="flex-1 overflow-y-auto p-4 space-y-4">
        {messages.map((msg, idx) => (
          <div key={idx} className={`flex gap-3 ${msg.role === 'user' ? 'flex-row-reverse' : ''}`}>
            <div className={`flex h-8 w-8 shrink-0 items-center justify-center rounded-full ${msg.role === 'user' ? 'bg-[#1f6fe5] text-white' : 'bg-[#eef4ff] text-[#1f6fe5]'}`}>
              {msg.role === 'user' ? <User size={16} /> : <Bot size={16} />}
            </div>
            <div className={`max-w-[85%] rounded-2xl px-4 py-3 text-sm leading-6 ${msg.role === 'user' ? 'rounded-tr-sm bg-[#1f6fe5] text-white' : 'rounded-tl-sm border border-[#e6ebf2] bg-[#fbfcfe] text-slate-600'}`}>
              {msg.text}
            </div>
          </div>
        ))}
        {isLoading && (
          <div className="flex gap-3">
            <div className="flex h-8 w-8 shrink-0 items-center justify-center rounded-full bg-[#eef4ff] text-[#1f6fe5]">
              <Bot size={16} />
            </div>
            <div className="flex items-center gap-2 rounded-2xl rounded-tl-sm border border-[#e6ebf2] bg-[#fbfcfe] px-4 py-3 text-slate-600">
              <Loader2 size={14} className="animate-spin text-[#1f6fe5]" />
              <span className="text-xs text-slate-400">Thinking...</span>
            </div>
          </div>
        )}
        <div ref={messagesEndRef} />
      </div>

      <div className="border-t border-[#eef2f7] bg-[#fbfcfe] p-4">
        <div className="relative">
          <input
            type="text"
            value={input}
            onChange={(e) => setInput(e.target.value)}
            onKeyDown={(e) => e.key === 'Enter' && handleSend()}
            placeholder="Ask Lumina to format, analyze..."
            className="w-full rounded-xl border border-[#e3e8ef] bg-white py-3 pl-4 pr-10 text-sm text-slate-700 placeholder:text-slate-400 transition-all focus:border-[#b6d0ff] focus:outline-none focus:ring-2 focus:ring-[#dce9ff]"
          />
          <button
            onClick={handleSend}
            disabled={!input.trim() || isLoading}
            className="absolute right-2 top-1/2 -translate-y-1/2 rounded-lg p-1.5 text-[#1f6fe5] transition-colors hover:bg-[#edf4ff] disabled:opacity-50 disabled:hover:bg-transparent"
          >
            <Send size={16} />
          </button>
        </div>
      </div>
    </div>
  );
};
