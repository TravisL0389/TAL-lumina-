
import React from 'react';
import { 
  Bold, 
  Italic, 
  Underline,
  Strikethrough,
  AlignLeft, 
  AlignCenter, 
  AlignRight, 
  PaintBucket,
  Eraser,
  Type as TypeIcon,
  Filter,
  ArrowUpDown,
  Wand2,
  BarChart,
  Database,
  TableProperties,
  Zap,
  LayoutGrid
} from 'lucide-react';
import { SheetState, CellId, CellData } from '../types';

interface ToolbarProps {
  sheet: SheetState;
  updateCell: (id: CellId, data: Partial<CellData>) => void;
  onGenerateContent: () => void;
  onOpenChart: () => void;
  onOpenImport: () => void;
  onFlashFill: () => void;
  onPivotTable: () => void;
  onCreateTable: () => void;
  onCleanData: () => void;
}

const Toolbar: React.FC<ToolbarProps> = ({ 
  sheet, 
  updateCell, 
  onGenerateContent,
  onOpenChart,
  onOpenImport,
  onFlashFill,
  onPivotTable,
  onCreateTable,
  onCleanData
}) => {
  const activeStyle = sheet.activeCell ? sheet.cells[sheet.activeCell]?.style || {} : {};

  const applyStyle = (style: React.CSSProperties) => {
    if (sheet.activeCell) {
      const currentStyle = sheet.cells[sheet.activeCell]?.style || {};
      
      // Toggle logic for certain styles
      const newStyle = { ...currentStyle };
      for (const [key, value] of Object.entries(style)) {
        if (newStyle[key as keyof React.CSSProperties] === value) {
          delete newStyle[key as keyof React.CSSProperties];
        } else {
          (newStyle as any)[key] = value;
        }
      }

      updateCell(sheet.activeCell, {
        style: newStyle
      });
    }
  };

  const clearFormatting = () => {
    if (sheet.activeCell) {
      updateCell(sheet.activeCell, { style: {} });
    }
  };

  return (
    <div className="toolbar-shell flex flex-wrap items-start gap-2 rounded-[1rem] border border-[#e6ebf2] bg-white px-3 py-2 text-slate-700">
      <div className="flex flex-wrap items-center gap-2 border-b border-[#eef2f7] pb-2 sm:border-b-0 sm:border-r sm:pb-0 sm:pr-3">
        <select 
          className="min-h-9 rounded-lg border border-[#e3e8ef] bg-[#fbfcfe] px-2 text-xs text-slate-700 outline-none transition-colors hover:bg-white"
          value={activeStyle.fontFamily || 'sans-serif'}
          onChange={(e) => applyStyle({ fontFamily: e.target.value })}
        >
          <option value="sans-serif">Sans Serif</option>
          <option value="serif">Serif</option>
          <option value="monospace">Monospace</option>
          <option value="Arial">Arial</option>
          <option value="Times New Roman">Times New Roman</option>
          <option value="Courier New">Courier New</option>
        </select>
        
        <select 
          className="min-h-9 rounded-lg border border-[#e3e8ef] bg-[#fbfcfe] px-2 text-xs text-slate-700 outline-none transition-colors hover:bg-white"
          value={activeStyle.fontSize || '11px'}
          onChange={(e) => applyStyle({ fontSize: e.target.value })}
        >
          <option value="9px">9</option>
          <option value="10px">10</option>
          <option value="11px">11</option>
          <option value="12px">12</option>
          <option value="14px">14</option>
          <option value="18px">18</option>
          <option value="24px">24</option>
        </select>
      </div>

      <div className="flex flex-wrap items-center gap-1 border-b border-[#eef2f7] pb-2 sm:border-b-0 sm:border-r sm:pb-0 sm:pr-3">
        <button 
          onClick={() => applyStyle({ fontWeight: 'bold' })}
          className={`min-h-9 min-w-9 rounded-lg transition-colors ${activeStyle.fontWeight === 'bold' ? 'bg-[#edf4ff] text-[#1f6fe5]' : 'text-slate-500 hover:bg-[#f3f6fb] hover:text-slate-800'}`}
        >
          <Bold className="w-4 h-4" />
        </button>
        <button 
          onClick={() => applyStyle({ fontStyle: 'italic' })}
          className={`min-h-9 min-w-9 rounded-lg transition-colors ${activeStyle.fontStyle === 'italic' ? 'bg-[#edf4ff] text-[#1f6fe5]' : 'text-slate-500 hover:bg-[#f3f6fb] hover:text-slate-800'}`}
        >
          <Italic className="w-4 h-4" />
        </button>
        <button 
          onClick={() => applyStyle({ textDecoration: 'underline' })}
          className={`min-h-9 min-w-9 rounded-lg transition-colors ${activeStyle.textDecoration === 'underline' ? 'bg-[#edf4ff] text-[#1f6fe5]' : 'text-slate-500 hover:bg-[#f3f6fb] hover:text-slate-800'}`}
        >
          <Underline className="w-4 h-4" />
        </button>
        <button 
          onClick={() => applyStyle({ textDecoration: 'line-through' })}
          className={`min-h-9 min-w-9 rounded-lg transition-colors ${activeStyle.textDecoration === 'line-through' ? 'bg-[#edf4ff] text-[#1f6fe5]' : 'text-slate-500 hover:bg-[#f3f6fb] hover:text-slate-800'}`}
        >
          <Strikethrough className="w-4 h-4" />
        </button>
      </div>

      <div className="flex flex-wrap items-center gap-1 border-b border-[#eef2f7] pb-2 sm:border-b-0 sm:border-r sm:pb-0 sm:pr-3">
        <div className="relative flex h-9 w-9 cursor-pointer items-center justify-center rounded-lg hover:bg-[#f3f6fb] group" title="Text Color">
          <TypeIcon className="w-4 h-4 text-slate-500" />
          <div className="absolute bottom-1 left-1.5 right-1.5 h-0.5 rounded-full" style={{ backgroundColor: activeStyle.color || '#f8fafc' }} />
          <input 
            type="color" 
            className="absolute inset-0 opacity-0 cursor-pointer"
            value={activeStyle.color || '#f8fafc'}
            onChange={(e) => applyStyle({ color: e.target.value })}
          />
        </div>
        <div className="relative flex h-9 w-9 cursor-pointer items-center justify-center rounded-lg hover:bg-[#f3f6fb] group" title="Fill Color">
          <PaintBucket className="w-4 h-4 text-slate-500" />
          <div className="absolute bottom-1 left-1.5 right-1.5 h-0.5 rounded-full" style={{ backgroundColor: activeStyle.backgroundColor || 'transparent' }} />
          <input 
            type="color" 
            className="absolute inset-0 opacity-0 cursor-pointer"
            value={activeStyle.backgroundColor || '#111925'}
            onChange={(e) => applyStyle({ backgroundColor: e.target.value })}
          />
        </div>
      </div>

      <div className="flex flex-wrap items-center gap-1 border-b border-[#eef2f7] pb-2 sm:border-b-0 sm:border-r sm:pb-0 sm:pr-3">
        <button 
          onClick={() => applyStyle({ textAlign: 'left' })}
          className={`min-h-9 min-w-9 rounded-lg transition-colors ${activeStyle.textAlign === 'left' ? 'bg-[#edf4ff] text-[#1f6fe5]' : 'text-slate-500 hover:bg-[#f3f6fb] hover:text-slate-800'}`}
        >
          <AlignLeft className="w-4 h-4" />
        </button>
        <button 
          onClick={() => applyStyle({ textAlign: 'center' })}
          className={`min-h-9 min-w-9 rounded-lg transition-colors ${activeStyle.textAlign === 'center' ? 'bg-[#edf4ff] text-[#1f6fe5]' : 'text-slate-500 hover:bg-[#f3f6fb] hover:text-slate-800'}`}
        >
          <AlignCenter className="w-4 h-4" />
        </button>
        <button 
          onClick={() => applyStyle({ textAlign: 'right' })}
          className={`min-h-9 min-w-9 rounded-lg transition-colors ${activeStyle.textAlign === 'right' ? 'bg-[#edf4ff] text-[#1f6fe5]' : 'text-slate-500 hover:bg-[#f3f6fb] hover:text-slate-800'}`}
        >
          <AlignRight className="w-4 h-4" />
        </button>
      </div>

      <div className="flex flex-wrap items-center gap-2 border-b border-[#eef2f7] pb-2 sm:border-b-0 sm:border-r sm:pb-0 sm:pr-3">
        <button 
          onClick={onGenerateContent}
          className="flex min-h-9 items-center gap-1.5 rounded-lg border border-[#cfe0ff] bg-[#f4f8ff] px-3 py-2 text-xs font-semibold text-[#1f6fe5] transition-colors hover:bg-[#eaf2ff]"
          title="AI Magic Fill"
        >
          <Wand2 className="w-3.5 h-3.5" />
          Magic Fill
        </button>
        <button 
          onClick={onFlashFill}
          className="flex min-h-9 items-center gap-1.5 rounded-lg border border-transparent px-3 py-2 text-xs font-semibold text-[#d97706] transition-colors hover:border-orange-200 hover:bg-orange-50"
          title="AI Flash Fill (Pattern Recognition)"
        >
          <Zap className="w-3.5 h-3.5" />
          Flash Fill
        </button>
        <button 
          onClick={onCleanData}
          className="flex min-h-9 items-center gap-1.5 rounded-lg border border-transparent px-3 py-2 text-xs font-semibold text-[#059669] transition-colors hover:border-emerald-200 hover:bg-emerald-50"
          title="AI Data Cleaning"
        >
          <Wand2 className="w-3.5 h-3.5" />
          Clean Data
        </button>
      </div>

      <div className="flex flex-wrap items-center gap-2 border-b border-[#eef2f7] pb-2 sm:border-b-0 sm:border-r sm:pb-0 sm:pr-3">
        <button 
          onClick={onOpenChart}
          className="flex min-h-9 items-center gap-1.5 rounded-lg px-3 py-2 text-xs font-medium text-slate-600 transition-colors hover:bg-[#f3f6fb]"
          title="Insert Chart"
        >
          <BarChart className="w-3.5 h-3.5" />
          Chart
        </button>
        <button 
          onClick={onCreateTable}
          className="flex min-h-9 items-center gap-1.5 rounded-lg px-3 py-2 text-xs font-medium text-slate-600 transition-colors hover:bg-[#f3f6fb]"
          title="Format as Table"
        >
          <LayoutGrid className="w-3.5 h-3.5" />
          Table
        </button>
        <button 
          onClick={onPivotTable}
          className="flex min-h-9 items-center gap-1.5 rounded-lg px-3 py-2 text-xs font-medium text-slate-600 transition-colors hover:bg-[#f3f6fb]"
          title="Pivot Table"
        >
          <TableProperties className="w-3.5 h-3.5" />
          Pivot
        </button>
        <button 
          onClick={onOpenImport}
          className="flex min-h-9 items-center gap-1.5 rounded-lg px-3 py-2 text-xs font-medium text-slate-600 transition-colors hover:bg-[#f3f6fb]"
          title="Power Query / Import"
        >
          <Database className="w-3.5 h-3.5" />
          Import
        </button>
      </div>

      <div className="flex flex-wrap items-center gap-2">
        <button 
          onClick={clearFormatting}
          className="flex min-h-9 items-center gap-1.5 rounded-lg px-3 py-2 text-xs font-medium text-rose-500 transition-colors hover:bg-rose-50"
          title="Clear Formatting"
        >
          <Eraser className="w-3.5 h-3.5" />
          Clear
        </button>
      </div>
    </div>
  );
};

export default Toolbar;
