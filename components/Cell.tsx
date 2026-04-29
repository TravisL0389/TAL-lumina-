
import React, { useState, useEffect, useRef, useMemo } from 'react';
import { CellId, CellData, SheetState, ConditionalRule, ValidationRule } from '../types';
import { evaluateAiFormula } from '../geminiService';
import { evaluateFormula } from '../utils/formulaEvaluator';
import { Sparkles, Loader2, AlertCircle } from 'lucide-react';

interface CellProps {
  id: CellId;
  data?: CellData;
  isActive: boolean;
  isSelected?: boolean;
  onClick: (e: React.MouseEvent) => void;
  updateCell: (id: CellId, data: Partial<CellData>) => void;
  sheet: SheetState;
  style?: React.CSSProperties;
}

const Cell: React.FC<CellProps> = ({ id, data, isActive, isSelected, onClick, updateCell, sheet, style }) => {
  const [isEditing, setIsEditing] = useState(false);
  const [localValue, setLocalValue] = useState(data?.formula || data?.value || '');
  const inputRef = useRef<HTMLDivElement>(null);
  const cellRef = useRef<HTMLDivElement>(null);

  useEffect(() => {
    if (!isEditing) {
      setLocalValue(data?.formula || data?.value || '');
    }
  }, [data, isEditing]);

  useEffect(() => {
    if (isActive) {
      if (isEditing) {
        if (inputRef.current) {
          inputRef.current.focus();
          // Move cursor to the end
          const range = document.createRange();
          const sel = window.getSelection();
          range.selectNodeContents(inputRef.current);
          range.collapse(false);
          sel?.removeAllRanges();
          sel?.addRange(range);
        }
      } else {
        cellRef.current?.focus();
      }
    }
  }, [isActive, isEditing]);

  const handleDoubleClick = () => {
    setIsEditing(true);
  };

  const handleBlur = (finalHTML?: string) => {
    const valToProcess = finalHTML !== undefined ? finalHTML : (inputRef.current?.innerHTML || localValue);
    processValue(valToProcess);
    setIsEditing(false);
  };

  const handleKeyDownDiv = (e: React.KeyboardEvent) => {
    if (isEditing) return;

    // Handle Delete/Backspace to clear
    if (e.key === 'Delete' || e.key === 'Backspace') {
      e.preventDefault();
      setLocalValue('');
      updateCell(id, { value: '', formula: '', displayValue: '' });
      return;
    }

    // Handle Enter to start editing
    if (e.key === 'Enter') {
      e.preventDefault();
      setIsEditing(true);
      return;
    }

    // Handle printable characters to start editing and overwrite
    if (e.key.length === 1 && !e.ctrlKey && !e.metaKey && !e.altKey) {
      setIsEditing(true);
      setLocalValue(e.key);
    }
  };

  const validateValue = (val: string, rule: ValidationRule): { isValid: boolean; error?: string } => {
    if (!val && rule.allowBlank) return { isValid: true };
    if (!val && !rule.allowBlank) return { isValid: false, error: rule.errorMessage };

    switch (rule.type) {
      case 'wholeNumber': {
        const num = parseInt(val);
        if (isNaN(num) || num.toString() !== val) return { isValid: false, error: rule.errorMessage };
        if (rule.operator === 'greaterThan' && rule.value1 && num <= parseInt(rule.value1)) return { isValid: false, error: rule.errorMessage };
        break;
      }
      case 'decimal': {
        const num = parseFloat(val);
        if (isNaN(num)) return { isValid: false, error: rule.errorMessage };
        break;
      }
      case 'date': {
        const date = new Date(val);
        if (isNaN(date.getTime())) return { isValid: false, error: rule.errorMessage };
        break;
      }
      case 'list': {
        if (rule.listItems && !rule.listItems.includes(val)) {
          return { isValid: false, error: rule.errorMessage };
        }
        break;
      }
    }
    return { isValid: true };
  };

  const processValue = async (htmlVal: string) => {
    const tempDiv = document.createElement('div');
    tempDiv.innerHTML = htmlVal;
    const plainText = tempDiv.textContent || tempDiv.innerText || '';
    
    let isValid = true;
    let validationError = "";

    const rule = sheet.validations[id];
    if (rule) {
      const v = validateValue(plainText, rule);
      isValid = v.isValid;
      validationError = v.error || "";
    }

    if (plainText.startsWith('=AI(')) {
      updateCell(id, { formula: plainText, displayValue: 'Thinking...', isLoading: true, isInvalid: !isValid, validationError });
      
      const prompt = plainText.slice(4, -1);
      const context = Object.entries(sheet.cells)
        .slice(0, 10)
        .map(([cid, c]) => {
          const tDiv = document.createElement('div');
          tDiv.innerHTML = (c as CellData).displayValue;
          return `${cid}: ${tDiv.textContent || tDiv.innerText || ''}`;
        })
        .join(", ");
        
      const result = await evaluateAiFormula(prompt, context);
      updateCell(id, { formula: plainText, displayValue: result, isLoading: false, isInvalid: !isValid, validationError });
    } else if (plainText.startsWith('=')) {
      const result = evaluateFormula(plainText, sheet);
      updateCell(id, { formula: plainText, displayValue: result, isLoading: false, isInvalid: !isValid, validationError });
    } else {
      updateCell(id, { formula: '', value: htmlVal, displayValue: htmlVal, isLoading: false, isInvalid: !isValid, validationError });
    }
  };

  const onKeyDown = (e: React.KeyboardEvent<HTMLDivElement>) => {
    if (e.key === 'Enter' && !e.shiftKey) {
      e.preventDefault();
      handleBlur(e.currentTarget.innerHTML);
    }
    if (e.key === 'Escape') {
      setLocalValue(data?.formula || data?.value || '');
      setIsEditing(false);
    }

    // Rich text shortcuts
    if (e.metaKey || e.ctrlKey) {
      if (e.key === 'b') {
        e.preventDefault();
        document.execCommand('bold', false);
      } else if (e.key === 'i') {
        e.preventDefault();
        document.execCommand('italic', false);
      } else if (e.key === 'u') {
        e.preventDefault();
        document.execCommand('underline', false);
      }
    }
  };

  const ruleStyle = useMemo(() => {
    const val = data?.displayValue || '';
    const numVal = parseFloat(val);
    const matchedRule = sheet.rules.find(rule => {
      if (rule.conditionType === 'greaterThan' && !isNaN(numVal)) {
        return numVal > parseFloat(rule.threshold || '0');
      }
      if (rule.conditionType === 'lessThan' && !isNaN(numVal)) {
        return numVal < parseFloat(rule.threshold || '0');
      }
      if (rule.conditionType === 'equals') {
        return val === rule.threshold;
      }
      if (rule.conditionType === 'contains') {
        return val.includes(rule.threshold || '');
      }
      return false;
    });
    return matchedRule ? matchedRule.style : {};
  }, [data?.displayValue, sheet.rules]);

  const validationRule = sheet.validations[id];
  const isDropdown = validationRule?.type === 'list' && validationRule.listItems && validationRule.listItems.length > 0;

  return (
    <div 
      ref={cellRef}
      tabIndex={0}
      onClick={onClick}
      onDoubleClick={handleDoubleClick}
      onKeyDown={handleKeyDownDiv}
      className={`
        sheet-cell relative flex items-center border-r border-b border-[#dfe6f0] font-medium text-[#24364e] transition-colors outline-none
        ${isActive ? 'z-10 bg-[#eef4ff] ring-2 ring-inset ring-[#2b78ff] shadow-[0_0_0_1px_rgba(43,120,255,0.24)]' : 'bg-[#ffffff] hover:bg-[#f8fbff]'}
        ${isSelected && !isActive ? 'bg-[#f1f6ff]' : ''}
        ${data?.isLoading ? 'bg-[#edf8ff]' : ''}
        ${data?.isInvalid ? 'bg-[#fff1f3]' : ''}
      `}
      title={data?.validationError}
      style={{ ...data?.style, ...ruleStyle, ...style }}
    >
      {isEditing ? (
        <div
          ref={inputRef}
          contentEditable
          suppressContentEditableWarning
          className="sheet-cell-editor absolute inset-0 z-20 h-full w-full overflow-auto whitespace-pre-wrap border-none bg-[#ffffff] font-mono text-[#17314c] shadow-2xl outline-none"
          onBlur={(e) => handleBlur(e.currentTarget.innerHTML)}
          onKeyDown={onKeyDown}
          dangerouslySetInnerHTML={{ __html: localValue }}
        />
      ) : isDropdown ? (
        <select
          className="h-full w-full appearance-none border-none bg-transparent text-[#24364e] outline-none cursor-pointer text-[inherit]"
          value={data?.displayValue || ''}
          onChange={(e) => processValue(e.target.value)}
        >
          <option value=""></option>
          {validationRule.listItems!.map((item, i) => (
            <option key={i} value={item}>{item}</option>
          ))}
        </select>
      ) : (
        <div 
          className={`pointer-events-none w-full overflow-hidden text-ellipsis whitespace-nowrap ${data?.isLoading ? 'text-[#2f8aa4]' : ''}`}
          dangerouslySetInnerHTML={{ __html: data?.displayValue || '' }}
        />
      )}
      
      {isDropdown && !isEditing && (
        <div className="absolute right-2 pointer-events-none">
          <svg className="w-3 h-3 text-[#94a3b8]" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M19 9l-7 7-7-7"></path></svg>
        </div>
      )}

      {data?.isLoading && (
        <div className="absolute right-1.5 flex items-center">
          <Loader2 className="w-3 h-3 animate-spin text-[#2b78ff]" />
        </div>
      )}

      {data?.isInvalid && !data?.isLoading && (
        <div className="absolute top-0 right-0 h-0 w-0 border-t-[6px] border-t-rose-400 border-l-[6px] border-l-transparent" />
      )}
      
      {data?.formula && !isEditing && !data?.isLoading && (
        <div className="absolute top-0 right-0 p-0.5">
           <Sparkles className="w-2.5 h-2.5 text-[#7bd5e1] opacity-60" />
        </div>
      )}
    </div>
  );
};

export default Cell;
