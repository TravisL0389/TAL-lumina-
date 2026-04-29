
import React from 'react';
import { AIInsight } from '../types';
import { 
  Lightbulb, 
  AlertCircle, 
  TrendingUp, 
  PieChart, 
  Sparkles,
  ChevronRight,
  Target,
  BarChart4,
  Info
} from 'lucide-react';

interface SidebarProps {
  insights: AIInsight[];
  summary: string;
  isAnalyzing: boolean;
  onAnalyze: () => void;
}

import InsightChart from './InsightChart';

const Sidebar: React.FC<SidebarProps> = ({ insights, summary, isAnalyzing, onAnalyze }) => {
  const getIcon = (type: string) => {
    switch (type) {
      case 'trend': return <TrendingUp className="w-4 h-4" />;
      case 'outlier': return <Target className="w-4 h-4" />;
      case 'correlation': return <BarChart4 className="w-4 h-4" />;
      case 'warning': return <AlertCircle className="w-4 h-4" />;
      case 'forecast': return <Lightbulb className="w-4 h-4" />;
      case 'risk': return <AlertCircle className="w-4 h-4" />;
      case 'opportunity': return <Sparkles className="w-4 h-4" />;
      default: return <Lightbulb className="w-4 h-4" />;
    }
  };

  const getColors = (type: string) => {
    switch (type) {
      case 'trend': return 'bg-indigo-100 text-indigo-700 border-indigo-200';
      case 'outlier': return 'bg-rose-100 text-rose-700 border-rose-200';
      case 'correlation': return 'bg-emerald-100 text-emerald-700 border-emerald-200';
      case 'warning': return 'bg-amber-100 text-amber-700 border-amber-200';
      case 'forecast': return 'bg-blue-100 text-blue-700 border-blue-200';
      case 'risk': return 'bg-orange-100 text-orange-700 border-orange-200';
      case 'opportunity': return 'bg-violet-100 text-violet-700 border-violet-200';
      default: return 'bg-slate-100 text-slate-700 border-slate-200';
    }
  };

  return (
    <aside className="w-80 bg-white border-l border-slate-200 flex flex-col shrink-0 overflow-y-auto z-20 shadow-2xl animate-in slide-in-from-right duration-500">
      <div className="p-5 border-b border-slate-100 flex items-center justify-between sticky top-0 bg-white/80 backdrop-blur-md z-10">
        <h2 className="text-sm font-black text-slate-800 flex items-center gap-2 tracking-tight">
          <Sparkles className="w-4 h-4 text-indigo-600" />
          INTELLIGENCE HUB
        </h2>
        <span className="text-[9px] font-black px-2 py-0.5 bg-indigo-600 text-white rounded-full uppercase tracking-tighter">AI 3.0</span>
      </div>

      <div className="flex-1 p-5 space-y-8">
        {isAnalyzing ? (
          <div className="flex flex-col items-center justify-center py-24 text-center">
            <div className="relative w-20 h-20 mb-6">
              <div className="absolute inset-0 border-4 border-indigo-50 rounded-full"></div>
              <div className="absolute inset-0 border-4 border-indigo-600 rounded-full border-t-transparent animate-spin"></div>
              <Sparkles className="absolute inset-0 m-auto w-8 h-8 text-indigo-600" />
            </div>
            <h3 className="text-base font-bold text-slate-800">Processing Complexities...</h3>
            <p className="text-xs text-slate-500 mt-2 px-6 font-medium leading-relaxed">
              Gemini is correlating indices and identifying statistical outliers in your dataset.
            </p>
          </div>
        ) : insights.length > 0 ? (
          <div className="space-y-6">
            {summary && (
              <div className="bg-slate-50 border border-slate-200 rounded-2xl p-4">
                <div className="flex items-center gap-2 mb-2 text-[10px] font-bold text-slate-400 uppercase tracking-widest">
                  <Info className="w-3 h-3" /> Data Summary
                </div>
                <p className="text-xs text-slate-700 leading-relaxed font-medium">{summary}</p>
              </div>
            )}

            <div className="space-y-4">
              <h4 className="text-[10px] font-bold text-slate-400 uppercase tracking-widest px-1">Key Findings</h4>
              {insights.map((insight, idx) => (
                <div key={idx} className="group p-4 bg-white rounded-2xl border border-slate-100 shadow-sm hover:shadow-md hover:border-indigo-200 transition-all cursor-default relative overflow-hidden">
                  <div className="absolute top-0 right-0 w-16 h-16 bg-gradient-to-br from-indigo-50/20 to-transparent pointer-events-none" />
                  <div className="flex items-start gap-3">
                    <div className={`p-2.5 rounded-xl border shrink-0 ${getColors(insight.type || '')}`}>
                      {getIcon(insight.type || '')}
                    </div>
                    <div className="flex-1">
                      <div className="flex justify-between items-start">
                        <h4 className="text-xs font-bold text-slate-800 mb-1 group-hover:text-indigo-700 transition-colors uppercase tracking-tight">{insight.title}</h4>
                      </div>
                      <p className="text-[11px] leading-relaxed text-slate-500 font-medium">{insight.description}</p>
                      
                      {/* Render Chart if visualization data exists */}
                      {insight.visualization && (
                        <InsightChart 
                          type={insight.visualization.chartType as any} 
                          data={insight.visualization.data}
                          xAxisLabel={insight.visualization.xAxisLabel}
                          yAxisLabel={insight.visualization.yAxisLabel}
                        />
                      )}

                      <div className="mt-3 flex items-center justify-between">
                         <span className="text-[9px] font-bold text-indigo-400 uppercase tracking-widest">{insight.type}</span>
                         <button className="flex items-center gap-1 text-[10px] font-bold text-indigo-600 hover:text-indigo-800 transition-colors">
                            Explore <ChevronRight className="w-3 h-3" />
                         </button>
                      </div>
                    </div>
                  </div>
                </div>
              ))}
            </div>
          </div>
        ) : (
          <div className="flex flex-col items-center justify-center py-20 text-center">
            <div className="w-16 h-16 bg-slate-50 rounded-3xl flex items-center justify-center mb-6">
               <PieChart className="w-8 h-8 text-slate-200" />
            </div>
            <h3 className="text-sm font-bold text-slate-400">Analysis Engine Ready</h3>
            <p className="text-[11px] text-slate-400 mt-2 px-10 leading-relaxed font-medium">
              Import or enter your data to unlock automatic trend detection and outlier analysis.
            </p>
            <button 
              onClick={onAnalyze}
              className="mt-8 px-8 py-3 bg-slate-900 text-white rounded-2xl text-xs font-black hover:bg-indigo-600 transition-all shadow-xl shadow-indigo-100 group"
            >
              <div className="flex items-center gap-2">
                 <Sparkles className="w-3.5 h-3.5 group-hover:rotate-12 transition-transform" />
                 INITIALIZE ANALYSIS
              </div>
            </button>
          </div>
        )}

        <div className="pt-8 border-t border-slate-100">
           <div className="bg-indigo-950 rounded-2xl p-5 text-white shadow-2xl relative overflow-hidden group">
              <div className="absolute top-0 right-0 w-32 h-32 bg-white/5 rounded-full -mr-16 -mt-16 group-hover:scale-110 transition-transform duration-700" />
              <h4 className="text-[10px] font-black uppercase tracking-widest text-indigo-400 mb-2">Pro Tip</h4>
              <p className="text-xs font-medium leading-relaxed mb-4 text-indigo-50">
                Type <code className="bg-white/10 px-1.5 py-0.5 rounded text-indigo-300">=AI("...")</code> in any cell to perform complex cross-row computations instantly.
              </p>
              <button className="w-full py-2 bg-indigo-500 rounded-xl text-[10px] font-black uppercase tracking-widest hover:bg-indigo-400 transition-colors">
                 Learn Syntax
              </button>
           </div>
        </div>
      </div>

      <div className="p-5 border-t border-slate-100 bg-slate-50/50">
        <div className="flex items-center gap-3">
           <div className="w-8 h-8 rounded-full bg-green-500/10 flex items-center justify-center">
              <div className="w-2 h-2 rounded-full bg-green-500 animate-pulse" />
           </div>
           <div>
              <p className="text-[10px] font-bold text-slate-700">Predictive Engine Live</p>
              <p className="text-[9px] font-medium text-slate-400 uppercase tracking-tighter">Low Latency Monitoring</p>
           </div>
        </div>
      </div>
    </aside>
  );
};

export default Sidebar;
