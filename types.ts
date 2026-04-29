
import React from 'react';

export type CellId = string; // e.g., "A1", "B10"

export interface CellData {
  value: string;
  formula: string;
  displayValue: string;
  style?: React.CSSProperties;
  isLoading?: boolean;
  isInvalid?: boolean;
  validationError?: string;
}

export interface MergeInfo {
  rowSpan: number;
  colSpan: number;
}

export interface ConditionalRule {
  id: string;
  description: string;
  conditionType: 'greaterThan' | 'lessThan' | 'equals' | 'contains' | 'custom';
  threshold?: string;
  style: React.CSSProperties;
}

export type ValidationType = 'any' | 'wholeNumber' | 'decimal' | 'date' | 'textLength' | 'custom' | 'list';
export type ValidationOperator = 'between' | 'notBetween' | 'equal' | 'notEqual' | 'greaterThan' | 'lessThan' | 'greaterThanOrEqual' | 'lessThanOrEqual';

export interface ValidationRule {
  type: ValidationType;
  operator?: ValidationOperator;
  value1?: string;
  value2?: string;
  formula?: string;
  listItems?: string[];
  allowBlank: boolean;
  errorMessage: string;
}

export interface TableStyle {
  id: string;
  startCell: CellId;
  endCell: CellId;
  headerStyle: React.CSSProperties;
  rowStyle: React.CSSProperties;
  altRowStyle: React.CSSProperties;
}

export interface SheetState {
  cells: Record<CellId, CellData>;
  activeCell: CellId | null;
  selectionStart: CellId | null;
  selectionEnd: CellId | null;
  rowCount: number;
  columnCount: number;
  rules: ConditionalRule[];
  validations: Record<CellId, ValidationRule>;
  merges: Record<CellId, MergeInfo>;
  tables: TableStyle[];
}

export interface AIInsight {
  title: string;
  description: string;
  type: 'info' | 'warning' | 'suggestion' | 'chart' | 'trend' | 'outlier' | 'correlation' | 'forecast' | 'risk' | 'opportunity';
  confidence?: number;
  visualization?: {
    chartType: 'bar' | 'line' | 'pie' | 'scatter';
    data: { label: string; value: number }[];
    xAxisLabel?: string;
    yAxisLabel?: string;
  };
}
