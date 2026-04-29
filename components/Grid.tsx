
import React, { useMemo, useState, useCallback, useEffect } from 'react';
import { SheetState, CellId, CellData, MergeInfo } from '../types';
import Cell from './Cell';

interface GridProps {
  sheet: SheetState;
  setActiveCell: (id: CellId, multi?: boolean) => void;
  updateCell: (id: CellId, data: Partial<CellData>) => void;
}

const Grid: React.FC<GridProps> = ({ sheet, setActiveCell, updateCell }) => {
  const [isDragging, setIsDragging] = useState(false);
  const getColumnLabel = useCallback((index: number) => {
    let label = '';
    let current = index + 1;
    while (current > 0) {
      const mod = (current - 1) % 26;
      label = String.fromCharCode(65 + mod) + label;
      current = Math.floor((current - mod) / 26);
    }
    return label;
  }, []);

  const columns = useMemo(() => {
    return Array.from({ length: sheet.columnCount }, (_, i) => getColumnLabel(i));
  }, [getColumnLabel, sheet.columnCount]);

  const rows = useMemo(() => {
    return Array.from({ length: sheet.rowCount }, (_, i) => i + 1);
  }, [sheet.rowCount]);

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

  const isCellHidden = (r: number, c: number) => {
    for (const [masterId, merge] of Object.entries(sheet.merges) as [string, MergeInfo][]) {
      const master = parseCellId(masterId);
      if (
        r >= master.r && 
        r < master.r + merge.rowSpan && 
        c >= master.c && 
        c < master.c + merge.colSpan
      ) {
        if (r === master.r && c === master.c) return false; // Master itself is not hidden
        return true;
      }
    }
    return false;
  };

  const isCellInSelection = (r: number, c: number) => {
    if (!sheet.selectionStart || !sheet.selectionEnd) return false;
    const s = parseCellId(sheet.selectionStart);
    const e = parseCellId(sheet.selectionEnd);
    const r1 = Math.min(s.r, e.r);
    const r2 = Math.max(s.r, e.r);
    const c1 = Math.min(s.c, e.c);
    const c2 = Math.max(s.c, e.c);
    return r >= r1 && r <= r2 && c >= c1 && c <= c2;
  };

  const handleMouseDown = (cellId: CellId, e: React.MouseEvent) => {
    if (e.button !== 0) return; // Only left click
    setIsDragging(true);
    setActiveCell(cellId, e.shiftKey);
  };

  const handleMouseEnter = (cellId: CellId) => {
    if (isDragging) {
      setActiveCell(cellId, true); // multi=true updates selectionEnd
    }
  };

  const handleMouseUp = () => {
    setIsDragging(false);
  };

  useEffect(() => {
    window.addEventListener('mouseup', handleMouseUp);
    return () => {
      window.removeEventListener('mouseup', handleMouseUp);
    };
  }, []);

  return (
    <div 
      className="sheet-grid grid select-none gap-[1px] border-b border-r border-[#dfe6f0] bg-[#dfe6f0]"
      style={{
        gridTemplateColumns: `var(--sheet-index-width) repeat(${sheet.columnCount}, var(--sheet-cell-width))`,
        gridTemplateRows: `var(--sheet-header-height) repeat(${sheet.rowCount}, var(--sheet-row-height))`,
        width: 'max-content'
      }}
      onMouseLeave={() => setIsDragging(false)}
    >
      {/* Corner */}
      <div className="sticky top-0 left-0 z-30 flex items-center justify-center border-r border-b border-[#dfe6f0] bg-[#f5f8fc]">
        <div className="w-1.5 h-1.5 rounded-full bg-[#b6c4d7]" />
      </div>

      {/* Col Headers */}
      {columns.map((col, idx) => (
        <div 
          key={col} 
          className="sticky top-0 z-20 flex items-center justify-center border-b border-[#dfe6f0] bg-[#f8fbff] text-[0.68rem] font-semibold text-[#6f8096] select-none"
        >
          {col}
        </div>
      ))}

      {/* Rows */}
      {rows.map(rowNum => {
        return (
          <React.Fragment key={rowNum}>
            {/* Row Header */}
            <div className="sticky left-0 z-20 flex items-center justify-center border-r border-[#dfe6f0] bg-[#f8fbff] text-[0.64rem] font-semibold text-[#7b8ea5] select-none">
              {rowNum}
            </div>

            {columns.map((col, colIdx) => {
              const c = colIdx + 1;
              const r = rowNum;
              const cellId = `${col}${rowNum}`;
              
              if (isCellHidden(r, c)) return null;

              const merge = sheet.merges[cellId];
              const wrapperStyle: React.CSSProperties = merge ? {
                gridRow: `span ${merge.rowSpan}`,
                gridColumn: `span ${merge.colSpan}`,
                height: 'auto',
                width: 'auto'
              } : {};

              let cellStyle: React.CSSProperties = { width: '100%', height: '100%' };

              // Apply table styles
              for (const table of sheet.tables) {
                const start = parseCellId(table.startCell);
                const end = parseCellId(table.endCell);
                const r1 = Math.min(start.r, end.r);
                const r2 = Math.max(start.r, end.r);
                const c1 = Math.min(start.c, end.c);
                const c2 = Math.max(start.c, end.c);

                if (r >= r1 && r <= r2 && c >= c1 && c <= c2) {
                  if (r === r1) {
                    cellStyle = { ...cellStyle, ...table.headerStyle };
                  } else if ((r - r1) % 2 === 0) {
                    cellStyle = { ...cellStyle, ...table.altRowStyle };
                  } else {
                    cellStyle = { ...cellStyle, ...table.rowStyle };
                  }
                }
              }

              return (
                <div
                  key={cellId}
                  onMouseDown={(e) => handleMouseDown(cellId, e)}
                  onMouseEnter={() => handleMouseEnter(cellId)}
                  style={wrapperStyle}
                  className="w-full h-full"
                >
                  <Cell 
                    id={cellId}
                    data={sheet.cells[cellId]}
                    isActive={sheet.activeCell === cellId}
                    isSelected={isCellInSelection(r, c)}
                    onClick={(e) => {}} // Handled by mousedown
                    updateCell={updateCell}
                    sheet={sheet}
                    style={cellStyle}
                  />
                </div>
              );
            })}
          </React.Fragment>
        );
      })}
    </div>
  );
};

export default Grid;
