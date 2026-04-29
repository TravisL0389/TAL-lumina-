import React, { Suspense, lazy, useCallback, useEffect, useMemo, useState } from 'react';
import Markdown from 'react-markdown';
import remarkGfm from 'remark-gfm';
import {
  AlertTriangle,
  ArrowUpRight,
  BarChart3,
  Bot,
  BrainCircuit,
  ChevronDown,
  ChevronLeft,
  ChevronRight,
  ChevronUp,
  Clock3,
  Combine,
  DatabaseZap,
  FolderKanban,
  FunctionSquare,
  HelpCircle,
  LayoutDashboard,
  Palette,
  PanelLeftClose,
  PanelLeftOpen,
  PlayCircle,
  Plus,
  Radar,
  Rows4,
  Search,
  ShieldCheck,
  Sigma,
  Save,
  Sparkles,
  Target,
  Undo2,
  Redo2,
  Wand2,
  Workflow,
  Zap,
  PencilLine,
  Check,
} from 'lucide-react';
import { CellData, CellId, SheetState, AIInsight, ConditionalRule, ValidationRule } from './types';
import Grid from './components/Grid';
import Toolbar from './components/Toolbar';
import FormulaBar from './components/FormulaBar';
const ConditionalFormattingModal = lazy(() => import('./components/ConditionalFormattingModal'));
const DataValidationModal = lazy(() => import('./components/DataValidationModal'));
const GenerateContentModal = lazy(() => import('./components/GenerateContentModal'));
const ChartModal = lazy(() => import('./components/ChartModal'));
const ImportModal = lazy(() => import('./components/ImportModal'));
const InsightChart = lazy(() => import('./components/InsightChart'));
const CopilotSidebar = lazy(() =>
  import('./components/CopilotSidebar').then((module) => ({ default: module.CopilotSidebar }))
);
import {
  analyzeSheetData,
  answerDataQuery,
  cleanDataRange,
  flashFillRange,
  generateCellContent,
} from './geminiService';

const LazyPanelFallback = () => (
  <div className="flex min-h-[180px] items-center justify-center rounded-3xl border border-white/10 bg-slate-950/80 p-6 text-sm text-slate-300 shadow-2xl shadow-black/30">
    Loading Lumina module...
  </div>
);

const LuminaMark: React.FC<{ className?: string }> = ({ className = 'h-8 w-8' }) => (
  <svg viewBox="0 0 40 40" fill="none" aria-hidden="true" className={className}>
    <g stroke="currentColor" strokeLinecap="round" strokeLinejoin="round">
      <path d="M20 4.5V12" strokeWidth="2.4" />
      <path d="M20 28V35.5" strokeWidth="2.4" />
      <path d="M4.5 20H12" strokeWidth="2.4" />
      <path d="M28 20H35.5" strokeWidth="2.4" />
      <path d="M9.2 9.2L14.6 14.6" strokeWidth="2.2" />
      <path d="M25.4 25.4L30.8 30.8" strokeWidth="2.2" />
      <path d="M9.2 30.8L14.6 25.4" strokeWidth="2.2" />
      <path d="M25.4 14.6L30.8 9.2" strokeWidth="2.2" />
      <path
        d="M20 10.8L23.1 16.9L29.2 20L23.1 23.1L20 29.2L16.9 23.1L10.8 20L16.9 16.9L20 10.8Z"
        fill="currentColor"
        strokeWidth="1.7"
      />
    </g>
  </svg>
);

const LuminaLogo: React.FC<{ showName?: boolean; compact?: boolean }> = ({ showName = true, compact = false }) => (
  <div className="flex items-center gap-2.5">
    <div
      className={`flex shrink-0 items-center justify-center rounded-[1.1rem] bg-[linear-gradient(180deg,#f8fbff_0%,#edf4ff_100%)] text-[#2d6df6] ${
        compact ? 'h-9 w-9' : 'h-10 w-10'
      }`}
    >
      <LuminaMark className={compact ? 'h-6 w-6' : 'h-6.5 w-6.5'} />
    </div>
    {showName && (
      <div className="leading-none">
        <div className="text-[1.05rem] font-bold tracking-[-0.03em] text-slate-900">Lumina Sheets</div>
      </div>
    )}
  </div>
);

const INITIAL_ROWS = 64;
const INITIAL_COLS = 18;
const SAVE_STORAGE_KEY = 'lumina-sheets-ai-saved-workbooks';

type ScenarioKey = 'conservative' | 'base' | 'aggressive';

interface WorkbookPreset {
  id: string;
  name: string;
  description: string;
  category: string;
  data: string[][];
  quickQuestions: string[];
  scenarios: Record<ScenarioKey, { revenue: string; margin: string; runway: string; risk: string }>;
  health: { label: string; detail: string; tone: 'stable' | 'watch' | 'critical' };
  activitySeed: Array<{ title: string; detail: string; tone: 'neutral' | 'positive' | 'watch' }>;
}

interface ActivityItem {
  id: string;
  title: string;
  detail: string;
  tone: 'neutral' | 'positive' | 'watch';
}

interface WorkbookDocument {
  id: string;
  name: string;
  sheet: SheetState;
  history: SheetState[];
  future: SheetState[];
  savedAt: string | null;
}

const WORKBOOKS: WorkbookPreset[] = [
  {
    id: 'revenue-command',
    name: 'Revenue Command',
    category: 'Executive revenue control',
    description:
      'An operating workbook built for ARR pacing, retention risk, ownership clarity, and board-ready forecasting.',
    data: [
      ['Region', 'Q1 ARR ($M)', 'Q2 ARR ($M)', 'Q3 ARR ($M)', 'Q4 ARR ($M)', 'Expansion %', 'Churn %', 'Forecast ARR ($M)', 'Risk', 'Owner'],
      ['North America', '4.2', '4.8', '5.1', '5.6', '17', '4', '20.4', 'Low', 'Maya'],
      ['EMEA', '2.7', '2.9', '3.2', '3.4', '11', '6', '12.2', 'Medium', 'Jon'],
      ['LATAM', '1.4', '1.8', '1.9', '2.1', '22', '8', '7.6', 'Watch', 'Sofia'],
      ['APAC', '3.0', '3.3', '3.8', '4.1', '15', '5', '14.5', 'Low', 'Noah'],
      ['Strategic', '5.4', '5.7', '6.3', '6.9', '19', '3', '24.3', 'Low', 'Rae'],
      ['Mid-Market', '2.3', '2.5', '2.6', '2.8', '9', '7', '10.2', 'Medium', 'Drew'],
      ['SMB', '1.2', '1.4', '1.7', '1.9', '25', '10', '6.2', 'Watch', 'Ari'],
    ],
    quickQuestions: [
      'Which region is slipping on retention?',
      'Summarize the ARR outlook for leadership.',
      'Where should we focus expansion plays first?',
      'Turn this into a board-ready narrative.',
    ],
    scenarios: {
      conservative: { revenue: '$89.4M', margin: '28.6%', runway: '16 months', risk: 'Retention drag in LATAM + SMB' },
      base: { revenue: '$94.8M', margin: '31.2%', runway: '18 months', risk: 'Healthy with two watch zones' },
      aggressive: { revenue: '$101.6M', margin: '34.1%', runway: '21 months', risk: 'Execution-heavy but attainable' },
    },
    health: { label: 'Stable Core', detail: 'Forecast confidence is high, but two segments need guided intervention.', tone: 'stable' },
    activitySeed: [
      { title: 'Q4 board pack draft ready', detail: 'Narrative packet can be generated directly from the current range.', tone: 'positive' },
      { title: 'Retention watch triggered', detail: 'LATAM and SMB churn bands crossed your watch threshold.', tone: 'watch' },
      { title: 'Expansion pattern detected', detail: 'Strategic accounts continue to outpace plan by 6.2%.', tone: 'positive' },
    ],
  },
  {
    id: 'pipeline-war-room',
    name: 'Pipeline War Room',
    category: 'Deal and forecast control',
    description:
      'A close-plan cockpit for sales leaders who need coverage, commit accuracy, deal risk, and next-step accountability in one place.',
    data: [
      ['Rep', 'Pipeline ($M)', 'Commit ($M)', 'Best Case ($M)', 'Win Rate %', 'Cycle Days', 'Coverage', 'Next Step', 'Risk'],
      ['Chloe', '3.8', '1.6', '2.4', '41', '34', '3.2x', 'Security review', 'Low'],
      ['Marcus', '2.9', '1.1', '1.9', '33', '46', '2.6x', 'Pricing alignment', 'Medium'],
      ['Talia', '4.6', '2.1', '3.1', '44', '29', '3.5x', 'Procurement signoff', 'Low'],
      ['Iris', '2.1', '0.7', '1.4', '28', '52', '1.9x', 'Champion rebuild', 'Watch'],
      ['Evan', '3.4', '1.5', '2.2', '39', '38', '2.8x', 'Legal redlines', 'Medium'],
      ['Nico', '1.7', '0.4', '0.9', '24', '57', '1.6x', 'Discovery gap', 'Watch'],
    ],
    quickQuestions: [
      'Which reps are most at risk this quarter?',
      'Create a close-plan summary I can send to leadership.',
      'Where is coverage weakest?',
      'Recommend three coaching actions for the team.',
    ],
    scenarios: {
      conservative: { revenue: '$17.8M', margin: '26.4%', runway: '13 months', risk: 'Coverage softness in two books' },
      base: { revenue: '$20.6M', margin: '29.0%', runway: '15 months', risk: 'Close plan achievable with targeted support' },
      aggressive: { revenue: '$23.4M', margin: '32.8%', runway: '18 months', risk: 'Requires commit uplift and cycle compression' },
    },
    health: { label: 'Watch List Active', detail: 'Forecast is recoverable, but bottom-quartile reps need intervention now.', tone: 'watch' },
    activitySeed: [
      { title: 'Deal slippage clustered', detail: 'Cycle times above 50 days now correlate with two watch-risk books.', tone: 'watch' },
      { title: 'Commit tier sharpened', detail: 'High-confidence deals are concentrated in three rep books.', tone: 'positive' },
      { title: 'Automation playbook available', detail: 'Generate a rep-by-rep close checklist in one click.', tone: 'neutral' },
    ],
  },
  {
    id: 'ops-cost-lens',
    name: 'Ops Cost Lens',
    category: 'Spend and efficiency control',
    description:
      'A planning model for budget variance, automation leverage, hiring pressure, and operational health across teams.',
    data: [
      ['Department', 'Budget ($M)', 'Actual ($M)', 'Variance %', 'Open Roles', 'Automation Hours', 'SLA', 'Health', 'Lead'],
      ['Support', '1.8', '1.9', '5.6', '3', '420', '94%', 'Watch', 'Harper'],
      ['Implementation', '2.4', '2.2', '-8.3', '1', '310', '97%', 'Strong', 'Lena'],
      ['Finance', '1.2', '1.1', '-6.8', '0', '180', '99%', 'Strong', 'Theo'],
      ['RevOps', '1.6', '1.9', '18.8', '2', '265', '92%', 'Watch', 'Cami'],
      ['People', '1.0', '1.1', '10.0', '4', '140', '96%', 'Moderate', 'June'],
      ['Security', '2.1', '2.0', '-4.8', '1', '390', '99%', 'Strong', 'Isaac'],
    ],
    quickQuestions: [
      'Where are we overspending the most?',
      'Turn this into an ops efficiency review.',
      'Which teams need automation investment first?',
      'Summarize the hiring and SLA story for leadership.',
    ],
    scenarios: {
      conservative: { revenue: '$42.1M saved', margin: '14.3%', runway: '11 months', risk: 'Hiring and SLA strain remain elevated' },
      base: { revenue: '$47.8M saved', margin: '18.2%', runway: '13 months', risk: 'Balanced, with RevOps and Support still under pressure' },
      aggressive: { revenue: '$53.6M saved', margin: '21.6%', runway: '16 months', risk: 'Depends on accelerated automation capture' },
    },
    health: { label: 'Efficiency Drift', detail: 'Automation leverage is solid, but labor pressure is compounding in two teams.', tone: 'critical' },
    activitySeed: [
      { title: 'Variance spike detected', detail: 'RevOps and People both exceeded budget tolerance this month.', tone: 'watch' },
      { title: 'SLA resilience holding', detail: 'Core delivery teams remain above the 95% target.', tone: 'positive' },
      { title: 'Automation upside found', detail: 'Support can reclaim another 60 hours per month with a workflow pass.', tone: 'neutral' },
    ],
  },
];

const PLAYBOOKS = [
  {
    id: 'board-brief',
    title: 'Executive Brief',
    detail: 'Turn the active sheet into an investor-grade narrative with highlights, concerns, and actions.',
    prompt: 'Create an executive summary with key drivers, watchouts, and actions from this spreadsheet.',
    icon: BrainCircuit,
  },
  {
    id: 'variance-watch',
    title: 'Variance Watch',
    detail: 'Run a smart sweep for anomalies, slipped trends, and hidden risk pockets.',
    prompt: 'Audit this spreadsheet for anomalies, variance spikes, and risk concentration. Be specific.',
    icon: Radar,
  },
  {
    id: 'workflow-pack',
    title: 'Automation Pack',
    detail: 'Generate next-step tasks, owner routing, and recurring review motions from the dataset.',
    prompt: 'Generate workflow recommendations and owner actions from this spreadsheet in a concise list.',
    icon: Workflow,
  },
];

const SOURCE_ITEMS = [
  { id: 'erp', label: 'ERP_Financials', icon: DatabaseZap, prompt: 'Summarize the most useful financial signals we should pull from ERP_Financials for this workbook.' },
  { id: 'crm', label: 'CRM_Deals', icon: FolderKanban, prompt: 'Explain how CRM_Deals should inform forecast risk and revenue planning in this workbook.' },
  { id: 'hr', label: 'HR_Planning', icon: Combine, prompt: 'Show how HR_Planning could be tied into headcount, spend, and capacity planning for this workbook.' },
  { id: 'market', label: 'Market_Data', icon: ShieldCheck, prompt: 'Suggest how Market_Data should be layered into the active workbook for better decision support.' },
];

const SIDEBAR_AUTOMATIONS = [
  {
    id: 'forecast-refresh',
    title: 'Monthly Forecast Refresh',
    detail: 'Last run 1h ago',
    prompt: 'Generate a monthly forecast refresh checklist for this workbook, including owner actions.',
  },
  {
    id: 'anomaly-alert',
    title: 'Revenue Anomaly Alert',
    detail: 'Last run 2h ago',
    prompt: 'Audit this workbook for revenue anomalies and explain the most important risk signal.',
  },
  {
    id: 'variance-commentary',
    title: 'Variance Commentary',
    detail: 'Last run 3h ago',
    prompt: 'Write concise variance commentary for the biggest changes in this workbook.',
  },
];

const toneClasses: Record<ActivityItem['tone'], string> = {
  neutral: 'text-slate-500 bg-slate-100',
  positive: 'text-emerald-600 bg-emerald-50',
  watch: 'text-orange-600 bg-orange-50',
};

const parseCellId = (id: string | null) => {
  if (!id) return null;
  const colMatch = id.match(/[A-Z]+/);
  const rowMatch = id.match(/[0-9]+/);
  if (!colMatch || !rowMatch) return null;
  let colIndex = 0;
  for (let i = 0; i < colMatch[0].length; i += 1) {
    colIndex = colIndex * 26 + (colMatch[0].charCodeAt(i) - 64);
  }
  return { r: Number(rowMatch[0]), c: colIndex };
};

const getCellIdFromCoords = (r: number, c: number) => {
  let col = '';
  let temp = c;
  while (temp > 0) {
    const mod = (temp - 1) % 26;
    col = String.fromCharCode(65 + mod) + col;
    temp = Math.floor((temp - mod) / 26);
  }
  return `${col}${r}`;
};

const getSelectionBounds = (selectionStart: string | null, selectionEnd: string | null) => {
  const start = parseCellId(selectionStart);
  const end = parseCellId(selectionEnd);
  if (!start || !end) return null;
  return {
    r1: Math.min(start.r, end.r),
    r2: Math.max(start.r, end.r),
    c1: Math.min(start.c, end.c),
    c2: Math.max(start.c, end.c),
  };
};

const cloneSheetState = (sheet: SheetState): SheetState =>
  typeof structuredClone === 'function' ? structuredClone(sheet) : JSON.parse(JSON.stringify(sheet));

const createBlankSheetState = (): SheetState => ({
  cells: {},
  activeCell: 'A1',
  selectionStart: 'A1',
  selectionEnd: 'A1',
  rowCount: INITIAL_ROWS,
  columnCount: INITIAL_COLS,
  rules: [],
  validations: {},
  merges: {},
  tables: [],
});

const createWorkbookDocument = (name: string, sheet: SheetState = createBlankSheetState()): WorkbookDocument => ({
  id: `doc_${Date.now()}_${Math.random().toString(36).slice(2, 8)}`,
  name,
  sheet,
  history: [],
  future: [],
  savedAt: null,
});

const INITIAL_DOCUMENT = createWorkbookDocument('Untitled Sheet 1');

const buildSheetFromMatrix = (data: string[][]): SheetState => {
  const usedCols = Math.max(...data.map((row) => row.length), 10);
  const columnCount = Math.max(usedCols, INITIAL_COLS);
  const rowCount = Math.max(data.length + 16, INITIAL_ROWS);
  const cells: Record<CellId, CellData> = {};

  data.forEach((row, rowIndex) => {
    row.forEach((value, colIndex) => {
      const cellId = getCellIdFromCoords(rowIndex + 1, colIndex + 1);
      const header = data[0]?.[colIndex] || '';
      const numericValue = Number(value);
      const isHeader = rowIndex === 0;

      cells[cellId] = {
        value,
        displayValue: value,
        formula: '',
        style: isHeader
          ? {
              fontWeight: 700,
              color: '#f8fbff',
              backgroundColor: '#1f6fe5',
              letterSpacing: '0.02em',
            }
          : !Number.isNaN(numericValue) && /(ARR|Budget|Actual|Variance|Hours|Rate|Coverage|Commit|Case|Pipeline)/i.test(header)
            ? { textAlign: 'right', color: '#23344d' }
            : { color: '#23344d' },
      };
    });
  });

  return {
    cells,
    activeCell: 'B2',
    selectionStart: 'B2',
    selectionEnd: 'E6',
    rowCount,
    columnCount,
    rules: [],
    validations: {},
    merges: {},
    tables: [
      {
        id: `table_${Date.now()}`,
        startCell: 'A1',
        endCell: getCellIdFromCoords(data.length, usedCols),
        headerStyle: {
          backgroundColor: '#1f6fe5',
          color: '#f8fbff',
          fontWeight: 'bold',
        },
        rowStyle: {
          backgroundColor: '#ffffff',
        },
        altRowStyle: {
          backgroundColor: '#f8fbff',
        },
      },
    ],
  };
};

const App: React.FC = () => {
  const [activeWorkbookId, setActiveWorkbookId] = useState(WORKBOOKS[0].id);
  const [selectedScenario, setSelectedScenario] = useState<ScenarioKey>('base');
  const [documents, setDocuments] = useState<WorkbookDocument[]>([INITIAL_DOCUMENT]);
  const [activeDocumentId, setActiveDocumentId] = useState(INITIAL_DOCUMENT.id);
  const [analysis, setAnalysis] = useState<{ summary: string; insights: AIInsight[] }>({ summary: '', insights: [] });
  const [activityLog, setActivityLog] = useState<ActivityItem[]>(
    WORKBOOKS[0].activitySeed.map((item, index) => ({ id: `seed-${index}`, ...item })),
  );
  const [showNavigator, setShowNavigator] = useState(true);
  const [showSidebar, setShowSidebar] = useState(true);
  const [showRulesModal, setShowRulesModal] = useState(false);
  const [showValidationModal, setShowValidationModal] = useState(false);
  const [showGenerateModal, setShowGenerateModal] = useState(false);
  const [showToolbar, setShowToolbar] = useState(true);
  const [showFormulaBar, setShowFormulaBar] = useState(true);
  const [showBottomPanel, setShowBottomPanel] = useState(true);
  const [showChartModal, setShowChartModal] = useState(false);
  const [showImportModal, setShowImportModal] = useState(false);
  const [showCopilot, setShowCopilot] = useState(false);
  const [globalQuery, setGlobalQuery] = useState('');
  const [documentNameDraft, setDocumentNameDraft] = useState(INITIAL_DOCUMENT.name);
  const [isRenamingDocument, setIsRenamingDocument] = useState(false);
  const [queryResult, setQueryResult] = useState<string | null>(
    'Ask Lumina to explain patterns, produce a briefing, clean a range, recommend actions, or simulate a plan. The responses can be narrative, tabular, or execution-oriented.',
  );
  const [isAnalyzing, setIsAnalyzing] = useState(false);
  const [isQuerying, setIsQuerying] = useState(false);

  const activeWorkbook = useMemo(
    () => WORKBOOKS.find((workbook) => workbook.id === activeWorkbookId) || WORKBOOKS[0],
    [activeWorkbookId],
  );
  const activeDocument = useMemo(
    () => documents.find((document) => document.id === activeDocumentId) || documents[0],
    [documents, activeDocumentId],
  );
  const sheet = activeDocument?.sheet || createBlankSheetState();
  const canUndo = Boolean(activeDocument?.history.length);
  const canRedo = Boolean(activeDocument?.future.length);
  const saveStatusLabel = activeDocument?.savedAt
    ? `Saved ${new Date(activeDocument.savedAt).toLocaleTimeString([], { hour: 'numeric', minute: '2-digit' })}`
    : 'Unsaved changes';

  const pushActivity = useCallback((title: string, detail: string, tone: ActivityItem['tone'] = 'neutral') => {
    setActivityLog((prev) => [
      { id: `${Date.now()}-${Math.random().toString(36).slice(2, 8)}`, title, detail, tone },
      ...prev,
    ].slice(0, 6));
  }, []);

  const applySheetUpdate = useCallback(
    (
      next: SheetState | ((prev: SheetState) => SheetState),
      options?: { recordHistory?: boolean; markDirty?: boolean },
    ) => {
      setDocuments((prev) =>
        prev.map((document) => {
          if (document.id !== activeDocumentId) return document;
          const nextSheet =
            typeof next === 'function'
              ? (next as (prev: SheetState) => SheetState)(cloneSheetState(document.sheet))
              : cloneSheetState(next);

          return {
            ...document,
            sheet: nextSheet,
            history:
              options?.recordHistory === false
                ? document.history
                : [...document.history.slice(-79), cloneSheetState(document.sheet)],
            future: options?.recordHistory === false ? document.future : [],
            savedAt: options?.markDirty === false ? document.savedAt : null,
          };
        }),
      );
    },
    [activeDocumentId],
  );

  const setSheet = useCallback(
    (next: SheetState | ((prev: SheetState) => SheetState)) => {
      applySheetUpdate(next);
    },
    [applySheetUpdate],
  );

  const setSheetTransient = useCallback(
    (next: SheetState | ((prev: SheetState) => SheetState)) => {
      applySheetUpdate(next, { recordHistory: false, markDirty: false });
    },
    [applySheetUpdate],
  );

  const getNextUntitledName = useCallback(
    (docs: WorkbookDocument[]) => {
      let index = 1;
      while (docs.some((document) => document.name === `Untitled Sheet ${index}`)) index += 1;
      return `Untitled Sheet ${index}`;
    },
    [],
  );

  const createNewWorkbook = useCallback(() => {
    const nextName = getNextUntitledName(documents);
    const nextDocument = createWorkbookDocument(nextName);
    setDocuments((prev) => [...prev, nextDocument]);
    setActiveDocumentId(nextDocument.id);
    setDocumentNameDraft(nextName);
    setIsRenamingDocument(true);
    setAnalysis({ summary: '', insights: [] });
    setQueryResult('A brand new workbook is ready. Start typing, import data, or ask Lumina to help build the sheet.');
    pushActivity('New workbook created', `${nextName} is ready for editing.`, 'positive');
  }, [documents, getNextUntitledName, pushActivity]);

  const commitDocumentName = useCallback(() => {
    const cleanedName = documentNameDraft.trim() || 'Untitled Sheet';
    setDocuments((prev) =>
      prev.map((document) =>
        document.id === activeDocumentId
          ? {
              ...document,
              name: cleanedName,
              savedAt: null,
            }
          : document,
      ),
    );
    setDocumentNameDraft(cleanedName);
    setIsRenamingDocument(false);
    pushActivity('Workbook renamed', `Current workbook renamed to ${cleanedName}.`, 'neutral');
  }, [activeDocumentId, documentNameDraft, pushActivity]);

  const handleUndo = useCallback(() => {
    if (!activeDocument?.history.length) return;
    setDocuments((prev) =>
      prev.map((document) => {
        if (document.id !== activeDocumentId || !document.history.length) return document;
        const previousSheet = document.history[document.history.length - 1];
        return {
          ...document,
          sheet: cloneSheetState(previousSheet),
          history: document.history.slice(0, -1),
          future: [cloneSheetState(document.sheet), ...document.future].slice(0, 80),
          savedAt: null,
        };
      }),
    );
  }, [activeDocument, activeDocumentId]);

  const handleRedo = useCallback(() => {
    if (!activeDocument?.future.length) return;
    setDocuments((prev) =>
      prev.map((document) => {
        if (document.id !== activeDocumentId || !document.future.length) return document;
        const [nextSheet, ...remainingFuture] = document.future;
        return {
          ...document,
          sheet: cloneSheetState(nextSheet),
          history: [...document.history.slice(-79), cloneSheetState(document.sheet)],
          future: remainingFuture,
          savedAt: null,
        };
      }),
    );
  }, [activeDocument, activeDocumentId]);

  const handleSave = useCallback(() => {
    if (!activeDocument) return;
    const now = new Date().toISOString();
    const payload = {
      id: activeDocument.id,
      name: activeDocument.name,
      savedAt: now,
      sheet: activeDocument.sheet,
    };
    window.localStorage.setItem(`${SAVE_STORAGE_KEY}:${activeDocument.id}`, JSON.stringify(payload));
    setDocuments((prev) =>
      prev.map((document) =>
        document.id === activeDocument.id
          ? {
              ...document,
              savedAt: now,
            }
          : document,
      ),
    );
    pushActivity('Workbook saved', `${activeDocument.name} was saved locally in your browser.`, 'positive');
  }, [activeDocument, pushActivity]);

  useEffect(() => {
    setAnalysis({ summary: '', insights: [] });
    setSelectedScenario('base');
    setActivityLog(activeWorkbook.activitySeed.map((item, index) => ({ id: `${activeWorkbook.id}-${index}`, ...item })));
    setQueryResult(`${activeWorkbook.description}\n\nTry one of the quick asks in the Decision Cockpit to turn this workbook into an operating system instead of a passive spreadsheet.`);
  }, [activeWorkbook]);

  useEffect(() => {
    if (activeDocument) {
      setDocumentNameDraft(activeDocument.name);
    }
  }, [activeDocument]);

  useEffect(() => {
    const handleKeydown = (event: KeyboardEvent) => {
      const isModifier = event.metaKey || event.ctrlKey;
      if (!isModifier) return;
      const key = event.key.toLowerCase();

      if (key === 's') {
        event.preventDefault();
        handleSave();
      } else if (key === 'z' && !event.shiftKey) {
        event.preventDefault();
        handleUndo();
      } else if (key === 'y' || (key === 'z' && event.shiftKey)) {
        event.preventDefault();
        handleRedo();
      }
    };

    window.addEventListener('keydown', handleKeydown);
    return () => window.removeEventListener('keydown', handleKeydown);
  }, [handleRedo, handleSave, handleUndo]);

  const getSheetAsCsv = useCallback(() => {
    const usedBounds = getSelectionBounds('A1', getCellIdFromCoords(Math.min(sheet.rowCount, 30), Math.min(sheet.columnCount, 12)));
    if (!usedBounds) return '';

    let csv = '';
    for (let r = usedBounds.r1; r <= usedBounds.r2; r += 1) {
      const row: string[] = [];
      for (let c = usedBounds.c1; c <= usedBounds.c2; c += 1) {
        row.push(sheet.cells[getCellIdFromCoords(r, c)]?.displayValue || '');
      }
      csv += `${row.join(',')}\n`;
    }
    return csv;
  }, [sheet]);

  const selectionSummary = useMemo(() => {
    const bounds = getSelectionBounds(sheet.selectionStart, sheet.selectionEnd);
    if (!bounds) {
      return { label: 'No range selected', count: 0, numericCount: 0, sum: 0, average: 0 };
    }

    const values: number[] = [];
    let count = 0;

    for (let r = bounds.r1; r <= bounds.r2; r += 1) {
      for (let c = bounds.c1; c <= bounds.c2; c += 1) {
        count += 1;
        const cellValue = sheet.cells[getCellIdFromCoords(r, c)]?.displayValue || '';
        const numericValue = Number(cellValue);
        if (!Number.isNaN(numericValue)) values.push(numericValue);
      }
    }

    return {
      label:
        sheet.selectionStart === sheet.selectionEnd
          ? sheet.selectionStart || 'Range'
          : `${sheet.selectionStart}:${sheet.selectionEnd}`,
      count,
      numericCount: values.length,
      sum: values.reduce((acc, value) => acc + value, 0),
      average: values.length ? values.reduce((acc, value) => acc + value, 0) / values.length : 0,
    };
  }, [sheet]);

  const workbookStats = useMemo(() => {
    const cells = Object.values(sheet.cells);
    const populated = cells.filter((cell) => Boolean(cell.displayValue || cell.value)).length;
    const formulas = cells.filter((cell) => Boolean(cell.formula)).length;
    const invalid = cells.filter((cell) => cell.isInvalid).length;
    const numeric = cells.filter((cell) => !Number.isNaN(Number(cell.displayValue || cell.value || ''))).length;
    return { populated, formulas, invalid, numeric };
  }, [sheet]);

  const healthMetrics = useMemo(() => {
    const csvRows = getSheetAsCsv()
      .trim()
      .split('\n')
      .map((row) => row.split(','));
    const dataRows = csvRows.slice(1);
    const blankFields = dataRows.reduce(
      (acc, row) => acc + row.filter((value) => value === '').length,
      0,
    );
    const watchSignals = dataRows.reduce(
      (acc, row) => acc + row.filter((value) => /watch|risk|medium/i.test(value)).length,
      0,
    );
    return {
      blankFields,
      watchSignals,
      density: csvRows.length > 1 ? Math.round((workbookStats.populated / (csvRows.length * (csvRows[0]?.length || 1))) * 100) : 0,
    };
  }, [getSheetAsCsv, workbookStats.populated]);

  const runGlobalQuery = useCallback(
    async (query: string) => {
      if (!query.trim()) return;
      setIsQuerying(true);
      setGlobalQuery(query);
      const answer = await answerDataQuery(query, getSheetAsCsv());
      setQueryResult(answer);
      setShowSidebar(true);
      setIsQuerying(false);
      pushActivity('AI query completed', query, 'neutral');
    },
    [getSheetAsCsv, pushActivity],
  );

  const handleAnalyze = async () => {
    setIsAnalyzing(true);
    const result = await analyzeSheetData(getSheetAsCsv());
    setAnalysis(result);
    setShowSidebar(true);
    setIsAnalyzing(false);
    pushActivity('Decision cockpit refreshed', 'Trend, outlier, and planning signals were regenerated.', 'positive');
  };

  const handleCleanData = async () => {
    setIsAnalyzing(true);
    const result = await cleanDataRange(getSheetAsCsv());

    if (result.cleanedCsv) {
      const rows = result.cleanedCsv
        .split('\n')
        .filter(Boolean)
        .map((row: string) => row.split(','));
      setSheet(buildSheetFromMatrix(rows));
    }

    setQueryResult(
      result.issues?.length
        ? `### Data Health Review\n\n- ${result.issues.join('\n- ')}`
        : '### Data Health Review\n\nNo major issues were detected. This workbook already looks structurally clean.',
    );
    setShowSidebar(true);
    setIsAnalyzing(false);
    pushActivity('Data quality pass finished', 'Lumina normalized and checked the active workbook.', 'positive');
  };

  const handleGenerateContent = async (prompt: string) => {
    if (!sheet.selectionStart || !sheet.selectionEnd) return;
    const bounds = getSelectionBounds(sheet.selectionStart, sheet.selectionEnd);
    if (!bounds) return;

    const targetCells: string[] = [];
    for (let r = bounds.r1; r <= bounds.r2; r += 1) {
      for (let c = bounds.c1; c <= bounds.c2; c += 1) {
        targetCells.push(getCellIdFromCoords(r, c));
      }
    }

    setSheet((prev) => {
      const nextCells = { ...prev.cells };
      targetCells.forEach((id) => {
        nextCells[id] = { ...nextCells[id], isLoading: true };
      });
      return { ...prev, cells: nextCells };
    });

    try {
      const contentMap = await generateCellContent(prompt, getSheetAsCsv(), targetCells);
      setSheet((prev) => {
        const nextCells = { ...prev.cells };
        targetCells.forEach((id) => {
          const value = String(contentMap[id] ?? nextCells[id]?.displayValue ?? '');
          nextCells[id] = {
            ...nextCells[id],
            value,
            displayValue: value,
            formula: value.startsWith('=') ? value : '',
            isLoading: false,
          };
        });
        return { ...prev, cells: nextCells };
      });
      pushActivity('Generated content inserted', `Applied AI content generation to ${targetCells.length} cells.`, 'positive');
    } catch (error) {
      setSheet((prev) => {
        const nextCells = { ...prev.cells };
        targetCells.forEach((id) => {
          nextCells[id] = { ...nextCells[id], isLoading: false };
        });
        return { ...prev, cells: nextCells };
      });
      pushActivity('Generation interrupted', 'The selected range could not be filled automatically.', 'watch');
    }
  };

  const handleImport = (data: string[][]) => {
    setSheet(buildSheetFromMatrix(data));
    setQueryResult('### New dataset loaded\n\nYour imported file is live in the workbook. Run an analysis or ask Lumina to clean, explain, or summarize it.');
    setShowSidebar(true);
    pushActivity('Dataset imported', `Loaded ${Math.max(data.length - 1, 0)} data rows into the active workbook.`, 'positive');
  };

  const handleFlashFill = async () => {
    if (!sheet.selectionStart || !sheet.selectionEnd) return;
    const bounds = getSelectionBounds(sheet.selectionStart, sheet.selectionEnd);
    if (!bounds) return;

    const targetCells: string[] = [];
    for (let r = bounds.r1; r <= bounds.r2; r += 1) {
      for (let c = bounds.c1; c <= bounds.c2; c += 1) {
        targetCells.push(getCellIdFromCoords(r, c));
      }
    }

    setSheet((prev) => {
      const nextCells = { ...prev.cells };
      targetCells.forEach((id) => {
        nextCells[id] = { ...nextCells[id], isLoading: true };
      });
      return { ...prev, cells: nextCells };
    });

    const prompt = `Fill ${sheet.selectionStart}:${sheet.selectionEnd} using the detected pattern from nearby spreadsheet data.`;
    const filledData = await flashFillRange(prompt, getSheetAsCsv());

    setSheet((prev) => {
      const nextCells = { ...prev.cells };
      targetCells.forEach((id) => {
        const value = String(filledData[id] ?? nextCells[id]?.displayValue ?? '');
        nextCells[id] = { ...nextCells[id], value, displayValue: value, isLoading: false };
      });
      return { ...prev, cells: nextCells };
    });
    pushActivity('Pattern fill applied', `Flash fill updated ${targetCells.length} cells.`, 'positive');
  };

  const handlePivotTable = async () => {
    setIsAnalyzing(true);
    const answer = await answerDataQuery(
      'Generate a pivot-style summary of the current spreadsheet with useful groupings and totals. Format it in markdown.',
      getSheetAsCsv(),
    );
    setQueryResult(answer);
    setShowSidebar(true);
    setIsAnalyzing(false);
    pushActivity('Pivot summary generated', 'Lumina created a grouped operating summary from the current sheet.', 'positive');
  };

  const handleCreateTable = () => {
    if (!sheet.selectionStart || !sheet.selectionEnd) return;
    setSheet((prev) => ({
      ...prev,
      tables: [
        ...prev.tables,
        {
          id: `table_${Date.now()}`,
          startCell: prev.selectionStart || 'A1',
          endCell: prev.selectionEnd || prev.selectionStart || 'A1',
          headerStyle: { backgroundColor: '#1f6fe5', color: '#f8fbff', fontWeight: 'bold' },
          rowStyle: { backgroundColor: '#ffffff', color: '#23344d' },
          altRowStyle: { backgroundColor: '#f8fbff', color: '#23344d' },
        },
      ],
    }));
    pushActivity('Structured table created', `Formatted ${sheet.selectionStart}:${sheet.selectionEnd} as an intelligent table region.`, 'neutral');
  };

  const updateCell = useCallback((id: CellId, data: Partial<CellData>) => {
    setSheet((prev) => ({
      ...prev,
      cells: {
        ...prev.cells,
        [id]: {
          ...(prev.cells[id] || { value: '', formula: '', displayValue: '' }),
          ...data,
        },
      },
    }));
  }, []);

  const addRule = (rule: ConditionalRule) => {
    setSheet((prev) => ({ ...prev, rules: [...prev.rules, rule] }));
    pushActivity('Guardrail added', `Conditional formatting rule "${rule.description}" is now active.`, 'neutral');
  };

  const removeRule = (id: string) => {
    setSheet((prev) => ({ ...prev, rules: prev.rules.filter((rule) => rule.id !== id) }));
  };

  const setValidation = (id: CellId, rule: ValidationRule | null) => {
    setSheet((prev) => {
      const nextValidations = { ...prev.validations };
      if (rule) nextValidations[id] = rule;
      else delete nextValidations[id];
      return { ...prev, validations: nextValidations };
    });
    pushActivity(
      rule ? 'Validation rule applied' : 'Validation rule removed',
      rule ? `Input protection is now active for ${id}.` : `Input protection was removed from ${id}.`,
      'neutral',
    );
  };

  const setActiveCell = useCallback((id: CellId, multi = false) => {
    setSheetTransient((prev) => (
      multi
        ? { ...prev, selectionEnd: id }
        : { ...prev, activeCell: id, selectionStart: id, selectionEnd: id }
    ));
  }, [setSheetTransient]);

  const handleMerge = () => {
    const bounds = getSelectionBounds(sheet.selectionStart, sheet.selectionEnd);
    if (!bounds || (sheet.selectionStart === sheet.selectionEnd)) return;
    const masterId = getCellIdFromCoords(bounds.r1, bounds.c1);

    setSheet((prev) => ({
      ...prev,
      merges: {
        ...prev.merges,
        [masterId]: { rowSpan: bounds.r2 - bounds.r1 + 1, colSpan: bounds.c2 - bounds.c1 + 1 },
      },
      activeCell: masterId,
      selectionStart: masterId,
      selectionEnd: masterId,
    }));
    pushActivity('Cells merged', `Created a unified view block at ${masterId}.`, 'neutral');
  };

  const handleUnmerge = () => {
    if (!sheet.activeCell) return;
    setSheet((prev) => {
      const nextMerges = { ...prev.merges };
      delete nextMerges[sheet.activeCell as string];
      return { ...prev, merges: nextMerges };
    });
    pushActivity('Cells unmerged', `Restored editable cells around ${sheet.activeCell}.`, 'neutral');
  };

  const isMerged = Boolean(sheet.activeCell && sheet.merges[sheet.activeCell]);

  const getChartData = () => {
    const bounds = getSelectionBounds(sheet.selectionStart, sheet.selectionEnd);
    if (!bounds || bounds.r2 - bounds.r1 < 1 || bounds.c2 - bounds.c1 < 1) return [];

    const headers: string[] = [];
    for (let c = bounds.c1 + 1; c <= bounds.c2; c += 1) {
      headers.push(sheet.cells[getCellIdFromCoords(bounds.r1, c)]?.displayValue || `Col ${c}`);
    }

    const data = [];
    for (let r = bounds.r1 + 1; r <= bounds.r2; r += 1) {
      const rowData: Record<string, number | string> = {
        name: sheet.cells[getCellIdFromCoords(r, bounds.c1)]?.displayValue || `Row ${r}`,
      };
      for (let c = bounds.c1 + 1; c <= bounds.c2; c += 1) {
        const value = Number(sheet.cells[getCellIdFromCoords(r, c)]?.displayValue || '0');
        rowData[headers[c - bounds.c1 - 1]] = Number.isNaN(value) ? 0 : value;
      }
      data.push(rowData);
    }
    return data;
  };

  const applyRangeValues = (range: string, values: string[][]) => {
    const [start] = range.split(':');
    const coords = parseCellId(start);
    if (!coords) return;

    setSheet((prev) => {
      const nextCells = { ...prev.cells };
      values.forEach((row, rowIndex) => {
        row.forEach((value, colIndex) => {
          const cellId = getCellIdFromCoords(coords.r + rowIndex, coords.c + colIndex);
          nextCells[cellId] = { ...nextCells[cellId], value, displayValue: value };
        });
      });
      return { ...prev, cells: nextCells };
    });
    pushActivity('Copilot action applied', `Updated range ${range} with generated content.`, 'positive');
  };

  const applyRangeFormatting = (range: string, style: React.CSSProperties) => {
    const [start, end] = range.split(':');
    const startCoords = parseCellId(start);
    const endCoords = parseCellId(end || start);
    if (!startCoords || !endCoords) return;

    setSheet((prev) => {
      const nextCells = { ...prev.cells };
      for (let r = startCoords.r; r <= endCoords.r; r += 1) {
        for (let c = startCoords.c; c <= endCoords.c; c += 1) {
          const cellId = getCellIdFromCoords(r, c);
          nextCells[cellId] = {
            ...nextCells[cellId],
            style: { ...(nextCells[cellId]?.style || {}), ...style },
          };
        }
      }
      return { ...prev, cells: nextCells };
    });
    pushActivity('Copilot formatting applied', `Applied formatting to ${range}.`, 'positive');
  };

  return (
    <div className="app-root flex min-h-screen w-full overflow-x-hidden text-slate-100">
      <div className="pointer-events-none absolute inset-0 bg-[linear-gradient(rgba(148,163,184,0.04)_1px,transparent_1px),linear-gradient(90deg,rgba(148,163,184,0.04)_1px,transparent_1px)] bg-[size:32px_32px] opacity-25" />

      <div className="relative flex min-h-screen w-full flex-col">
        <header className="border-b border-[#e7edf5] bg-white px-4 py-3 md:px-6">
          <div className="flex flex-col gap-3 xl:flex-row xl:items-center xl:justify-between">
            <div className="flex items-center gap-3">
              <button
                onClick={() => setShowNavigator((prev) => !prev)}
                className="inline-flex rounded-xl border border-[#e3e8ef] bg-[#fbfcfe] p-2 text-slate-500 transition hover:bg-white"
                title="Toggle workspace navigator"
              >
                {showNavigator ? <PanelLeftClose className="h-4 w-4" /> : <PanelLeftOpen className="h-4 w-4" />}
              </button>
              <LuminaLogo compact />
              <div>
                <div className="flex items-center gap-2">
                  {isRenamingDocument ? (
                    <div className="flex items-center gap-2">
                      <input
                        value={documentNameDraft}
                        onChange={(event) => setDocumentNameDraft(event.target.value)}
                        onBlur={commitDocumentName}
                        onKeyDown={(event) => {
                          if (event.key === 'Enter') commitDocumentName();
                          if (event.key === 'Escape') {
                            setDocumentNameDraft(activeDocument?.name || 'Untitled Sheet');
                            setIsRenamingDocument(false);
                          }
                        }}
                        autoFocus
                        className="rounded-lg border border-[#dbe4f0] bg-white px-3 py-1 text-lg font-bold tracking-tight text-slate-900 outline-none focus:border-[#b6d0ff] focus:ring-2 focus:ring-[#dce9ff]"
                      />
                      <button
                        onClick={commitDocumentName}
                        className="inline-flex h-8 w-8 items-center justify-center rounded-lg bg-[#1f6fe5] text-white transition hover:bg-[#1b63cf]"
                        title="Save workbook name"
                      >
                        <Check className="h-4 w-4" />
                      </button>
                    </div>
                  ) : (
                    <>
                      <h1 className="text-lg font-bold tracking-tight text-slate-900">{activeDocument?.name || 'Untitled Sheet'}</h1>
                      <button
                        onClick={() => {
                          setDocumentNameDraft(activeDocument?.name || 'Untitled Sheet');
                          setIsRenamingDocument(true);
                        }}
                        className="inline-flex h-7 w-7 items-center justify-center rounded-lg border border-[#e3e8ef] bg-[#fbfcfe] text-slate-400 transition hover:bg-white hover:text-slate-700"
                        title="Rename workbook"
                      >
                        <PencilLine className="h-3.5 w-3.5" />
                      </button>
                    </>
                  )}
                </div>
                <p className="text-sm text-slate-500">
                  {saveStatusLabel} • {documents.length} open workbook{documents.length === 1 ? '' : 's'}
                </p>
              </div>
            </div>

            <div className="flex flex-1 flex-col gap-3 xl:mx-8 xl:max-w-3xl">
              <form
                onSubmit={(event) => {
                  event.preventDefault();
                  void runGlobalQuery(globalQuery);
                }}
                className="group relative"
              >
                <div className="absolute inset-y-0 left-4 flex items-center">
                  <Search className={`h-4 w-4 transition ${isQuerying ? 'animate-pulse text-[#1f6fe5]' : 'text-slate-400 group-focus-within:text-[#1f6fe5]'}`} />
                </div>
                <input
                  type="text"
                  className="h-11 w-full rounded-xl border border-[#e3e8ef] bg-[#fbfcfe] pl-11 pr-32 text-sm text-slate-700 outline-none transition placeholder:text-slate-400 focus:border-[#b6d0ff] focus:bg-white focus:ring-2 focus:ring-[#dce9ff] md:pr-36"
                  placeholder="Search sheets, data, and actions"
                  value={globalQuery}
                  onChange={(event) => setGlobalQuery(event.target.value)}
                />
                <button
                  type="submit"
                  disabled={isQuerying || !globalQuery.trim()}
                  className="absolute right-2 top-1.5 inline-flex h-8 items-center gap-2 rounded-lg bg-[#1f6fe5] px-3 text-xs font-semibold text-white transition hover:bg-[#1b63cf] disabled:cursor-not-allowed disabled:opacity-60"
                >
                  <Sparkles className="h-3.5 w-3.5" />
                  Ask Lumina
                </button>
              </form>

              <div className="hidden flex-wrap gap-2">
                {activeWorkbook.quickQuestions.map((question) => (
                  <button
                    key={question}
                    onClick={() => void runGlobalQuery(question)}
                    className="rounded-full border border-sky-100/14 bg-white/8 px-3 py-1.5 text-left text-xs font-semibold text-slate-100 transition hover:border-sky-300/35 hover:bg-sky-300/10 hover:text-white"
                  >
                    {question}
                  </button>
                ))}
              </div>
            </div>

            <div className="flex flex-wrap items-center gap-2">
              <button
                onClick={createNewWorkbook}
                className="inline-flex shrink-0 items-center gap-2 rounded-xl border border-[#dbe4f0] bg-[#f5f9ff] px-3 py-2 text-xs font-semibold text-[#1f6fe5] transition hover:bg-[#edf4ff]"
              >
                <Plus className="h-4 w-4" />
                New
              </button>
              <button
                onClick={handleUndo}
                disabled={!canUndo}
                className="inline-flex shrink-0 items-center gap-2 rounded-xl border border-[#e3e8ef] bg-white px-3 py-2 text-xs font-semibold text-slate-700 transition hover:bg-[#f8fbff] disabled:cursor-not-allowed disabled:opacity-45"
              >
                <Undo2 className="h-4 w-4" />
                Undo
              </button>
              <button
                onClick={handleRedo}
                disabled={!canRedo}
                className="inline-flex shrink-0 items-center gap-2 rounded-xl border border-[#e3e8ef] bg-white px-3 py-2 text-xs font-semibold text-slate-700 transition hover:bg-[#f8fbff] disabled:cursor-not-allowed disabled:opacity-45"
              >
                <Redo2 className="h-4 w-4" />
                Redo
              </button>
              <button
                onClick={handleSave}
                className="inline-flex shrink-0 items-center gap-2 rounded-xl border border-[#e3e8ef] bg-white px-3 py-2 text-xs font-semibold text-slate-700 transition hover:bg-[#f8fbff]"
              >
                <Save className="h-4 w-4" />
                Save
              </button>
              <button
                onClick={() => setShowCopilot(true)}
                className="inline-flex shrink-0 items-center gap-2 rounded-xl border border-[#dbe4f0] bg-[#f5f9ff] px-3 py-2 text-xs font-semibold text-[#1f6fe5] transition hover:bg-[#edf4ff]"
              >
                <Sparkles className="h-4 w-4" />
                Ask Lumina
              </button>
              <button
                onClick={handleAnalyze}
                className="inline-flex shrink-0 items-center gap-2 rounded-xl border border-[#e3e8ef] bg-white px-3 py-2 text-xs font-semibold text-slate-700 transition hover:bg-[#f8fbff]"
              >
                <BarChart3 className="h-4 w-4" />
                Analyze
              </button>
              <button
                onClick={() => void runGlobalQuery('Generate workflow recommendations and owner actions from this spreadsheet in a concise list.')}
                className="inline-flex shrink-0 items-center gap-2 rounded-xl border border-[#e3e8ef] bg-white px-3 py-2 text-xs font-semibold text-slate-700 transition hover:bg-[#f8fbff]"
              >
                <Workflow className="h-4 w-4" />
                Automate
              </button>
              <button
                onClick={() => setShowFormulaBar((prev) => !prev)}
                className="inline-flex shrink-0 items-center gap-2 rounded-xl border border-[#e3e8ef] bg-white px-3 py-2 text-xs font-semibold text-slate-700 transition hover:bg-[#f8fbff]"
              >
                <FunctionSquare className="h-4 w-4" />
                Model
              </button>
              <button
                onClick={() => setShowSidebar((prev) => !prev)}
                className="inline-flex shrink-0 items-center gap-2 rounded-xl bg-[#1f6fe5] px-3 py-2 text-xs font-semibold text-white transition hover:bg-[#1b63cf]"
              >
                <ArrowUpRight className="h-4 w-4" />
                Share
              </button>
            </div>
          </div>
        </header>

        <div className="workspace-body flex flex-1 flex-col xl:min-h-0 xl:flex-row">
          {showNavigator && (
            <aside className="hidden xl:flex xl:w-[15.5rem] xl:shrink-0 xl:flex-col xl:border-r xl:border-[#e7edf5] xl:bg-white xl:px-4 xl:py-4">
              <div className="mb-4 flex items-center justify-between gap-3 px-2">
                <LuminaLogo />
                <button
                  onClick={() => setShowNavigator(false)}
                  className="inline-flex rounded-xl border border-[#e3e8ef] bg-[#fbfcfe] p-2 text-slate-400 transition hover:bg-white hover:text-slate-700"
                  title="Hide workspace navigator"
                >
                  <PanelLeftClose className="h-4 w-4" />
                </button>
              </div>

              <button
                onClick={createNewWorkbook}
                className="mb-3 inline-flex items-center justify-between rounded-xl bg-[#1f6fe5] px-4 py-3 text-sm font-semibold text-white shadow-[0_12px_20px_rgba(31,111,229,0.18)]"
              >
                <span>+ New</span>
                <ChevronRight className="h-4 w-4 opacity-80" />
              </button>

              <div className="mb-5 rounded-xl border border-[#e7edf5] bg-[#fbfcfe] px-3 py-2 text-sm text-slate-400">
                Search workbook
              </div>

              <div className="mb-5">
                <div className="mb-3 flex items-center justify-between px-1">
                  <p className="text-[11px] font-semibold uppercase tracking-[0.24em] text-slate-400">Workbook</p>
                </div>
                <div className="space-y-1.5">
                  {documents.map((document) => (
                    <button
                      key={document.id}
                      onClick={() => setActiveDocumentId(document.id)}
                      className={`w-full rounded-xl border px-3 py-2.5 text-left transition ${
                        document.id === activeDocumentId
                          ? 'border-[#d7e5ff] bg-[#f2f7ff] shadow-[0_8px_18px_rgba(148,163,184,0.12)]'
                          : 'border-transparent bg-white hover:border-[#e7edf5] hover:bg-[#fbfcfe]'
                      }`}
                    >
                      <div className="flex items-center justify-between gap-3">
                        <div className="min-w-0">
                          <p className={`text-sm font-medium ${document.id === activeDocumentId ? 'text-[#1f6fe5]' : 'text-slate-700'}`}>{document.name}</p>
                          <p className="mt-1 text-xs text-slate-400">
                            {document.savedAt ? 'Saved locally' : 'Unsaved'}
                          </p>
                        </div>
                        <div className="flex items-center gap-2">
                          {document.id === activeDocumentId && <span className="h-2.5 w-2.5 rounded-full bg-[#3b82f6]" />}
                          <ChevronRight className="h-4 w-4 text-slate-500" />
                        </div>
                      </div>
                    </button>
                  ))}
                </div>
              </div>

              <div className="mb-5">
                <div className="mb-3 flex items-center justify-between px-1">
                  <p className="text-[11px] font-semibold uppercase tracking-[0.24em] text-slate-400">Sources</p>
                  <button className="text-slate-400 transition hover:text-slate-700" title="Add source">
                    <Plus className="h-4 w-4" />
                  </button>
                </div>
                <div className="space-y-3">
                  {SOURCE_ITEMS.map((source) => {
                    const Icon = source.icon;
                    return (
                      <button
                        key={source.id}
                        onClick={() => void runGlobalQuery(source.prompt)}
                        className="flex w-full items-center justify-between gap-3 rounded-xl border border-[#eef2f7] bg-[#fbfcfe] p-3 text-left transition hover:border-[#d7e5ff] hover:bg-[#f7faff]"
                      >
                        <div className="flex min-w-0 items-center gap-3">
                          <div className="rounded-xl bg-[#eef4ff] p-2 text-[#1f6fe5]">
                            <Icon className="h-4 w-4" />
                          </div>
                          <div className="min-w-0">
                            <p className="truncate text-sm font-medium text-slate-700">{source.label}</p>
                          </div>
                        </div>
                        <span className="inline-flex h-4 w-4 items-center justify-center rounded-md border border-emerald-200 bg-emerald-50 text-emerald-500">
                          <Check className="h-2.5 w-2.5" />
                        </span>
                      </button>
                    );
                  })}
                </div>
              </div>

              <div className="mb-5">
                <div className="mb-3 flex items-center justify-between px-1">
                  <p className="text-[11px] font-semibold uppercase tracking-[0.24em] text-slate-400">Automations</p>
                  <button className="text-slate-400 transition hover:text-slate-700" title="Add automation">
                    <Plus className="h-4 w-4" />
                  </button>
                </div>
                <div className="space-y-2.5">
                  {SIDEBAR_AUTOMATIONS.map((automation) => (
                    <button
                      key={automation.id}
                      onClick={() => void runGlobalQuery(automation.prompt)}
                      className="flex w-full items-start justify-between gap-3 rounded-xl border border-[#eef2f7] bg-white px-3 py-2.5 text-left transition hover:border-[#d7e5ff] hover:bg-[#f8fbff]"
                    >
                      <div className="min-w-0">
                        <p className="text-sm font-medium text-slate-700">{automation.title}</p>
                        <p className="mt-1 text-xs text-slate-400">{automation.detail}</p>
                      </div>
                      <span className="mt-1 h-2.5 w-2.5 shrink-0 rounded-full bg-emerald-400" />
                    </button>
                  ))}
                </div>
                <button className="mt-3 inline-flex w-full items-center justify-center gap-2 rounded-xl border border-[#e7edf5] bg-[#fbfcfe] px-4 py-3 text-sm font-medium text-slate-600 transition hover:bg-white">
                  <Plus className="h-4 w-4" />
                  New automation
                </button>
              </div>

              <div className="mt-auto rounded-[1.4rem] border border-[#eef2f7] bg-[#fbfcfe] p-4">
                <p className="text-[11px] font-semibold uppercase tracking-[0.24em] text-slate-400">Data Health</p>
                <div className="mt-4 flex items-center gap-3">
                  <div className="relative flex h-16 w-16 items-center justify-center rounded-full border-[6px] border-emerald-100 text-emerald-500">
                    <div className="absolute inset-[5px] rounded-full border-[4px] border-emerald-400 border-t-emerald-300" />
                    <span className="relative text-lg font-bold">92</span>
                  </div>
                  <div>
                    <p className="text-lg font-semibold text-emerald-500">Excellent</p>
                    <p className="text-xs text-slate-400">All checks passed</p>
                  </div>
                </div>
                <button className="mt-4 text-sm font-medium text-[#1f6fe5] transition hover:text-[#1b63cf]">View details</button>
              </div>
            </aside>
          )}

          <main className="workspace-main-scroll relative flex min-w-0 flex-1 flex-col overflow-x-hidden px-4 py-4 md:px-6">
            {!showNavigator && (
              <button
                onClick={() => setShowNavigator(true)}
                className="workspace-edge-tab workspace-edge-tab-left"
                title="Show workspace navigator"
              >
                <PanelLeftOpen className="h-4 w-4" />
                <span>Show Navigator</span>
              </button>
            )}

            {!showSidebar && (
              <button
                onClick={() => setShowSidebar(true)}
                className="workspace-edge-tab workspace-edge-tab-right"
                title="Show decision cockpit"
              >
                <span>Show Cockpit</span>
                <ChevronLeft className="h-4 w-4" />
              </button>
            )}

            {showNavigator && (
              <section className="mb-4 grid gap-4 xl:hidden">
                <div className="rounded-[28px] border border-[#e7edf5] bg-white p-4 shadow-[0_18px_38px_rgba(46,87,128,0.08)]">
                  <div className="flex flex-col gap-4 md:flex-row md:items-start md:justify-between">
                    <div className="max-w-2xl">
                      <p className="text-[11px] font-bold uppercase tracking-[0.24em] text-slate-400">Workbook focus</p>
                      <h2 className="mt-2 text-lg font-bold text-slate-900">{activeWorkbook.name}</h2>
                      <p className="mt-2 text-sm leading-6 text-slate-500">{activeWorkbook.description}</p>
                    </div>
                    <div className="min-w-0 rounded-2xl border border-[#e7edf5] bg-[#fbfcfe] p-3 md:max-w-xs">
                      <div className="flex items-center justify-between gap-3">
                        <div>
                          <p className="text-[11px] font-bold uppercase tracking-[0.24em] text-slate-400">Health</p>
                          <p className="mt-1 text-sm font-semibold text-slate-900">{activeWorkbook.health.label}</p>
                        </div>
                        <span
                          className={`rounded-full px-2 py-1 text-[10px] font-bold uppercase tracking-[0.18em] ${
                            activeWorkbook.health.tone === 'stable'
                              ? 'bg-emerald-50 text-emerald-600'
                              : activeWorkbook.health.tone === 'watch'
                                ? 'bg-orange-50 text-orange-600'
                                : 'bg-rose-50 text-rose-600'
                          }`}
                        >
                          {activeWorkbook.health.tone}
                        </span>
                      </div>
                      <p className="mt-2 text-xs leading-5 text-slate-400">{activeWorkbook.health.detail}</p>
                    </div>
                  </div>
                </div>

                <div className="grid gap-3 sm:grid-cols-2">
                  {PLAYBOOKS.map((playbook) => {
                    const Icon = playbook.icon;
                    return (
                      <button
                        key={playbook.id}
                        onClick={() => void runGlobalQuery(playbook.prompt)}
                        className="rounded-[26px] border border-[#e7edf5] bg-white p-4 text-left transition hover:border-[#d7e5ff] hover:bg-[#f8fbff]"
                      >
                        <div className="flex items-start gap-3">
                          <div className="rounded-2xl bg-[#eef4ff] p-2 text-[#1f6fe5]">
                            <Icon className="h-4 w-4" />
                          </div>
                          <div>
                            <p className="text-sm font-semibold text-slate-900">{playbook.title}</p>
                            <p className="mt-1 text-xs leading-5 text-slate-400">{playbook.detail}</p>
                          </div>
                        </div>
                      </button>
                    );
                  })}
                </div>
              </section>
            )}

            <div className="mb-4 flex gap-3 overflow-x-auto pb-1 xl:hidden no-scrollbar">
              {documents.map((document) => (
                <button
                  key={document.id}
                  onClick={() => setActiveDocumentId(document.id)}
                  className={`shrink-0 rounded-2xl border px-4 py-2 text-left transition ${
                    document.id === activeDocumentId
                      ? 'border-[#1f6fe5] bg-[#eef4ff] text-[#1f6fe5]'
                      : 'border-[#e3e8ef] bg-white text-slate-700'
                  }`}
                >
                  <div className="text-sm font-semibold">{document.name}</div>
                  <div className="text-xs text-slate-400">{document.savedAt ? 'Saved locally' : 'Unsaved'}</div>
                </button>
              ))}
              <button
                onClick={createNewWorkbook}
                className="shrink-0 rounded-2xl border border-dashed border-[#cfe0ff] bg-[#f5f9ff] px-4 py-2 text-sm font-semibold text-[#1f6fe5]"
              >
                + New
              </button>
            </div>

            <section className="hidden mb-4 grid gap-3 sm:grid-cols-2 2xl:grid-cols-4">
              <MetricCard
                icon={<LayoutDashboard className="h-4 w-4" />}
                eyebrow="Selection"
                title={selectionSummary.label}
                detail={`${selectionSummary.count} cells · ${selectionSummary.numericCount} numeric`}
                accent="from-[#1f4566]/55 to-transparent"
              />
              <MetricCard
                icon={<Sigma className="h-4 w-4" />}
                eyebrow="Range Math"
                title={selectionSummary.numericCount ? selectionSummary.sum.toFixed(2) : 'No numeric values'}
                detail={selectionSummary.numericCount ? `Average ${selectionSummary.average.toFixed(2)}` : 'Select a measurable range'}
                accent="from-[#1e4f47]/55 to-transparent"
              />
              <MetricCard
                icon={<Rows4 className="h-4 w-4" />}
                eyebrow="Data Density"
                title={`${workbookStats.populated} populated cells`}
                detail={`${workbookStats.numeric} numeric · ${workbookStats.formulas} formulas`}
                accent="from-[#304468]/55 to-transparent"
              />
              <MetricCard
                icon={<AlertTriangle className="h-4 w-4" />}
                eyebrow="Quality Signal"
                title={workbookStats.invalid ? `${workbookStats.invalid} invalid cells` : `${healthMetrics.watchSignals} watch signals`}
                detail={workbookStats.invalid ? 'Validation needs attention' : `${healthMetrics.blankFields} blanks · ${healthMetrics.density}% density`}
                accent="from-[#5a4721]/55 to-transparent"
              />
            </section>

            <section className="mb-4 flex items-center gap-2 overflow-x-auto pb-1 no-scrollbar">
              {documents.map((document) => (
                <button
                  key={document.id}
                  onClick={() => setActiveDocumentId(document.id)}
                  className={`shrink-0 rounded-t-xl border border-b-0 px-4 py-2 text-sm font-semibold transition ${
                    document.id === activeDocumentId
                      ? 'border-[#d7e5ff] bg-white text-[#1f6fe5] shadow-[0_-4px_14px_rgba(148,163,184,0.08)]'
                      : 'border-[#edf2f7] bg-[#f7faff] text-slate-500 hover:bg-white hover:text-slate-700'
                  }`}
                >
                  {document.name}
                </button>
              ))}
              <button
                onClick={createNewWorkbook}
                className="shrink-0 rounded-t-xl border border-dashed border-[#cfe0ff] bg-[#f5f9ff] px-4 py-2 text-sm font-semibold text-[#1f6fe5]"
              >
                + New sheet
              </button>
            </section>

            <div className="workspace-primary-grid mb-4">
              <section className="flex min-h-[26rem] flex-col overflow-hidden rounded-[1.4rem] border border-[#e7edf5] bg-white shadow-[0_20px_42px_rgba(148,163,184,0.12)] lg:min-h-[30rem]">
                <div className="border-b border-[#edf2f7] px-4 py-3 md:px-5">
                  <div className="flex flex-col gap-3 2xl:flex-row 2xl:items-center 2xl:justify-between">
                    <div>
                      <p className="text-[11px] font-semibold uppercase tracking-[0.24em] text-slate-400">Active workbook</p>
                      <h2 className="mt-1 text-xl font-bold text-slate-900">{activeDocument?.name || 'Untitled Sheet'}</h2>
                      <p className="mt-1 text-sm text-slate-500">{saveStatusLabel}</p>
                    </div>
                    <div className="flex flex-wrap gap-2">
                      <button
                        onClick={() => setShowToolbar((prev) => !prev)}
                        className="inline-flex items-center gap-2 rounded-xl border border-[#e3e8ef] bg-[#fbfcfe] px-4 py-2 text-xs font-semibold text-slate-700 transition hover:bg-white"
                      >
                        {showToolbar ? <ChevronUp className="h-4 w-4" /> : <ChevronDown className="h-4 w-4" />}
                        {showToolbar ? 'Hide Toolbar' : 'Show Toolbar'}
                      </button>
                      <button
                        onClick={handleAnalyze}
                        disabled={isAnalyzing}
                        className="inline-flex items-center gap-2 rounded-xl bg-[#1f6fe5] px-4 py-2 text-xs font-semibold text-white transition hover:bg-[#1b63cf] disabled:opacity-60"
                      >
                        <Sparkles className="h-4 w-4" />
                        Refresh Intelligence
                      </button>
                      <button
                        onClick={() => setShowImportModal(true)}
                        className="inline-flex items-center gap-2 rounded-xl border border-[#e3e8ef] bg-white px-4 py-2 text-xs font-semibold text-slate-700 transition hover:bg-[#f8fbff]"
                      >
                        <DatabaseZap className="h-4 w-4" />
                        Import Data
                      </button>
                      <button
                        onClick={() => void runGlobalQuery('Draft the next meeting agenda from this workbook, including decisions and owner follow-ups.')}
                        className="inline-flex items-center gap-2 rounded-xl border border-[#e3e8ef] bg-white px-4 py-2 text-xs font-semibold text-slate-700 transition hover:bg-[#f8fbff]"
                      >
                        <PlayCircle className="h-4 w-4" />
                        Build Review Pack
                      </button>
                      {!showSidebar && (
                        <button
                          onClick={() => setShowSidebar(true)}
                          className="inline-flex items-center gap-2 rounded-xl border border-[#e3e8ef] bg-[#fbfcfe] px-4 py-2 text-xs font-semibold text-slate-700 transition hover:bg-white"
                        >
                          <BarChart3 className="h-4 w-4" />
                          Show Decision Cockpit
                        </button>
                      )}
                    </div>
                  </div>
                </div>

                <div className="border-b border-[#edf2f7] px-3 py-2.5 md:px-4">
                  <div className="flex flex-wrap items-center justify-between gap-3">
                    <div>
                      <p className="text-[11px] font-semibold uppercase tracking-[0.24em] text-slate-400">Toolbar</p>
                      <p className="text-xs text-slate-500">Keep the spreadsheet chrome open, or tuck it away for more canvas room.</p>
                    </div>
                    <button
                      onClick={() => setShowToolbar((prev) => !prev)}
                      className="inline-flex items-center gap-2 rounded-full border border-[#dbe4f0] bg-[#f5f9ff] px-3 py-1.5 text-xs font-semibold text-[#1f6fe5] transition hover:bg-[#edf4ff]"
                    >
                      {showToolbar ? <ChevronUp className="h-3.5 w-3.5" /> : <ChevronDown className="h-3.5 w-3.5" />}
                      {showToolbar ? 'Pull up' : 'Pull down'}
                    </button>
                  </div>

                  {showToolbar ? (
                    <div className="mt-3">
                      <Toolbar
                        sheet={sheet}
                        updateCell={updateCell}
                        onGenerateContent={() => setShowGenerateModal(true)}
                        onOpenChart={() => setShowChartModal(true)}
                        onOpenImport={() => setShowImportModal(true)}
                        onFlashFill={handleFlashFill}
                        onPivotTable={handlePivotTable}
                        onCreateTable={handleCreateTable}
                        onCleanData={handleCleanData}
                      />
                    </div>
                  ) : (
                    <div className="mt-3 rounded-2xl border border-dashed border-[#d7e5ff] bg-[#f7faff] px-4 py-3 text-sm text-slate-500">
                      Toolbar hidden. Pull it back down whenever you need formatting, import, chart, or AI sheet tools.
                    </div>
                  )}
                </div>

                {showFormulaBar && (
                  <div className="border-b border-sky-100/10 px-3 py-3 md:px-4">
                    <FormulaBar
                      activeCell={sheet.activeCell}
                      cellData={sheet.activeCell ? sheet.cells[sheet.activeCell] || null : null}
                      updateCell={updateCell}
                      sheet={sheet}
                    />
                  </div>
                )}

                <div className="flex-1 overflow-hidden px-3 pb-3 md:px-4 md:pb-4">
                  <div className="grid-surface sheet-grid-frame h-full overflow-auto rounded-[1.2rem] border border-[#edf2f7] bg-[#fbfcfe] p-2">
                    <Grid sheet={sheet} setActiveCell={setActiveCell} updateCell={updateCell} />
                  </div>
                </div>
              </section>

              {showSidebar && (
                <aside className="flex min-h-[20rem] flex-col overflow-hidden rounded-[1.4rem] border border-[#e7edf5] bg-white shadow-[0_20px_42px_rgba(148,163,184,0.12)]">
                  <div className="border-b border-[#edf2f7] px-5 py-4">
                    <div className="flex items-center justify-between">
                      <div>
                        <p className="text-[11px] font-semibold uppercase tracking-[0.24em] text-slate-400">Insights</p>
                        <h3 className="mt-1 text-lg font-bold text-slate-900">Lumina Intelligence</h3>
                      </div>
                      <button
                        onClick={() => setShowSidebar(false)}
                        className="rounded-xl border border-[#e3e8ef] bg-white p-2 text-slate-400 transition hover:bg-[#f8fbff]"
                        title="Hide decision cockpit"
                      >
                        <ChevronRight className="h-4 w-4" />
                      </button>
                    </div>

                    <div className="mt-4 grid gap-3">
                      <div className="rounded-2xl border border-[#edf2f7] bg-[#fbfcfe] p-3">
                        <div className="flex items-center justify-between">
                          <div>
                            <p className="text-[11px] font-semibold uppercase tracking-[0.22em] text-slate-400">Revenue Forecast</p>
                            <p className="mt-1 text-sm font-semibold text-slate-900">{activeWorkbook.scenarios[selectedScenario].revenue}</p>
                          </div>
                          <div className="rounded-2xl bg-[#edf4ff] p-2 text-[#1f6fe5]">
                            <Target className="h-4 w-4" />
                          </div>
                        </div>
                        <div className="mt-3 flex gap-2">
                          {(['conservative', 'base', 'aggressive'] as ScenarioKey[]).map((scenario) => (
                            <button
                              key={scenario}
                              onClick={() => setSelectedScenario(scenario)}
                              className={`rounded-full px-3 py-1.5 text-[11px] font-bold capitalize transition ${
                                selectedScenario === scenario
                                  ? 'bg-[#1f6fe5] text-white'
                                  : 'bg-white text-slate-500 hover:bg-[#f5f9ff]'
                              }`}
                            >
                              {scenario}
                            </button>
                          ))}
                        </div>
                        <div className="mt-3 grid gap-2 text-xs text-slate-500">
                          <div className="flex items-center justify-between">
                            <span>Margin</span>
                            <span>{activeWorkbook.scenarios[selectedScenario].margin}</span>
                          </div>
                          <div className="flex items-center justify-between">
                            <span>Runway</span>
                            <span>{activeWorkbook.scenarios[selectedScenario].runway}</span>
                          </div>
                          <div className="rounded-2xl border border-[#edf2f7] bg-white px-3 py-2 text-slate-500">
                            {activeWorkbook.scenarios[selectedScenario].risk}
                          </div>
                        </div>
                      </div>

                      <div className="rounded-2xl border border-[#edf2f7] bg-[#fbfcfe] p-3">
                        <div className="mb-2 flex items-center gap-2 text-[11px] font-bold uppercase tracking-[0.22em] text-slate-400">
                          <Wand2 className="h-3.5 w-3.5 text-[#1f6fe5]" />
                          Quick asks
                        </div>
                        <div className="flex flex-wrap gap-2">
                          {activeWorkbook.quickQuestions.slice(0, 3).map((question) => (
                            <button
                              key={question}
                              onClick={() => void runGlobalQuery(question)}
                              className="rounded-full border border-[#e3e8ef] bg-white px-3 py-1.5 text-[11px] font-semibold text-slate-600 transition hover:border-[#d7e5ff] hover:bg-[#f5f9ff]"
                            >
                              {question}
                            </button>
                          ))}
                        </div>
                      </div>
                    </div>
                  </div>

                  <div className="flex-1 space-y-4 overflow-y-auto px-5 py-4">
                    <div className="rounded-3xl border border-[#edf2f7] bg-[#fbfcfe] p-4">
                      <div className="mb-3 flex items-center gap-2 text-[11px] font-bold uppercase tracking-[0.22em] text-slate-400">
                        <Sparkles className="h-3.5 w-3.5 text-[#1f6fe5]" />
                        View analysis
                      </div>
                      <div className="prose prose-sm max-w-none text-slate-600 prose-headings:text-slate-900 prose-strong:text-slate-900 prose-a:text-[#1f6fe5] prose-table:text-slate-600 prose-th:border-[#e7edf5] prose-th:bg-[#f5f9ff] prose-td:border-[#e7edf5] prose-code:text-[#1f6fe5]">
                        <Markdown remarkPlugins={[remarkGfm]}>{queryResult || 'Ask a question to turn the current workbook into decisions and next steps.'}</Markdown>
                      </div>
                    </div>

                    <div className="rounded-3xl border border-[#edf2f7] bg-[#fbfcfe] p-4">
                      <div className="mb-3 flex items-center gap-2 text-[11px] font-bold uppercase tracking-[0.22em] text-slate-400">
                        <Radar className="h-3.5 w-3.5 text-[#1f6fe5]" />
                        Data Health
                      </div>
                      <div className="grid gap-3">
                        <HealthRow label="Blank fields" value={healthMetrics.blankFields.toString()} />
                        <HealthRow label="Watch signals" value={healthMetrics.watchSignals.toString()} />
                        <HealthRow label="Sheet density" value={`${healthMetrics.density}%`} />
                      </div>
                    </div>

                    <div className="rounded-3xl border border-[#edf2f7] bg-[#fbfcfe] p-4">
                      <div className="mb-3 flex items-center gap-2 text-[11px] font-bold uppercase tracking-[0.22em] text-slate-400">
                        <ArrowUpRight className="h-3.5 w-3.5 text-[#1f6fe5]" />
                        Intelligence feed
                      </div>
                      {isAnalyzing ? (
                        <div className="rounded-2xl border border-[#dbe9ff] bg-[#f4f8ff] p-4 text-sm text-[#1f6fe5]">
                          Lumina is recalculating trends, correlations, and planning signals...
                        </div>
                      ) : analysis.insights.length ? (
                        <div className="space-y-3">
                          {analysis.summary && (
                            <div className="rounded-2xl border border-[#edf2f7] bg-white p-3 text-sm leading-6 text-slate-600">
                              {analysis.summary}
                            </div>
                          )}
                          {analysis.insights.slice(0, 4).map((insight, index) => (
                            <div key={`${insight.title}-${index}`} className="rounded-2xl border border-[#edf2f7] bg-white p-3">
                              <div className="flex items-start justify-between gap-3">
                                <div>
                                  <p className="text-sm font-semibold text-slate-900">{insight.title}</p>
                                  <p className="mt-1 text-xs leading-5 text-slate-400">{insight.description}</p>
                                </div>
                                <span className="rounded-full bg-[#edf4ff] px-2 py-1 text-[10px] font-semibold uppercase tracking-[0.18em] text-[#1f6fe5]">
                                  {insight.type}
                                </span>
                              </div>
                              {insight.visualization && (
                                <div className="mt-3 rounded-2xl border border-[#edf2f7] bg-white p-2">
                                  <InsightChart
                                    type={insight.visualization.chartType}
                                    data={insight.visualization.data}
                                    xAxisLabel={insight.visualization.xAxisLabel}
                                    yAxisLabel={insight.visualization.yAxisLabel}
                                  />
                                </div>
                              )}
                            </div>
                          ))}
                        </div>
                      ) : (
                        <div className="rounded-2xl border border-dashed border-[#d9e1eb] bg-white p-4 text-sm leading-6 text-slate-400">
                          Run a fresh analysis to populate trend calls, risk pockets, and suggested actions directly from the workbook.
                        </div>
                      )}
                    </div>
                  </div>
                </aside>
              )}

            </div>

            <div className="mb-3 flex justify-end">
              <button
                onClick={() => setShowBottomPanel((prev) => !prev)}
                className="inline-flex items-center gap-2 rounded-full border border-[#dbe4f0] bg-[#f5f9ff] px-3 py-1.5 text-xs font-semibold text-[#1f6fe5] transition hover:bg-[#edf4ff]"
              >
                {showBottomPanel ? <ChevronDown className="h-3.5 w-3.5" /> : <ChevronUp className="h-3.5 w-3.5" />}
                {showBottomPanel ? 'Hide Bottom Dock' : 'Show Bottom Dock'}
              </button>
            </div>

            {showBottomPanel ? (
              <section className="workspace-bottom-grid pb-4">
                <div className="rounded-[1.4rem] border border-[#e7edf5] bg-white p-5 shadow-[0_20px_42px_rgba(148,163,184,0.12)]">
                  <div className="mb-4 flex items-center justify-between">
                    <div>
                      <p className="text-[11px] font-semibold uppercase tracking-[0.24em] text-slate-400">Scenario Simulator</p>
                      <h3 className="mt-1 text-lg font-bold text-slate-900">Plan levers and modeled outcomes</h3>
                    </div>
                    <div className="rounded-2xl bg-[#edf4ff] p-2 text-[#1f6fe5]">
                      <Zap className="h-4 w-4" />
                    </div>
                  </div>
                  <div className="grid gap-3 sm:grid-cols-2">
                    <FeatureCard
                      icon={<BrainCircuit className="h-4 w-4" />}
                      title="AI Formula Coach"
                      detail="Explain, predict, and generate formulas from intent so teams do not need specialist spreadsheet syntax to move fast."
                    />
                    <FeatureCard
                      icon={<Target className="h-4 w-4" />}
                      title="Scenario Lab"
                      detail="Flip between conservative, base, and aggressive outcomes with context that reads like a planning partner."
                    />
                    <FeatureCard
                      icon={<Workflow className="h-4 w-4" />}
                      title="Automation Playbooks"
                      detail="Translate raw data into recurring workflows, owner actions, and review packs rather than static rows."
                    />
                    <FeatureCard
                      icon={<Radar className="h-4 w-4" />}
                      title="Decision Cockpit"
                      detail="Turn questions into guidance, summaries, anomalies, and visual evidence without exporting to another tool."
                    />
                  </div>
                </div>

                <div className="rounded-[1.4rem] border border-[#e7edf5] bg-white p-5 shadow-[0_20px_42px_rgba(148,163,184,0.12)]">
                  <div className="mb-4 flex items-center justify-between">
                    <div>
                      <p className="text-[11px] font-semibold uppercase tracking-[0.24em] text-slate-400">Activity Timeline</p>
                      <h3 className="mt-1 text-lg font-bold text-slate-900">Recent workbook actions</h3>
                    </div>
                    <div className="rounded-2xl bg-[#edf4ff] p-2 text-[#1f6fe5]">
                      <Clock3 className="h-4 w-4" />
                    </div>
                  </div>
                  <div className="space-y-3">
                    {activityLog.map((item) => (
                      <div key={item.id} className="rounded-2xl border border-[#edf2f7] bg-[#fbfcfe] p-3">
                        <div className="flex items-start gap-3">
                          <div className={`mt-0.5 rounded-full px-2 py-1 text-[10px] font-bold uppercase tracking-[0.18em] ${toneClasses[item.tone]}`}>
                            {item.tone}
                          </div>
                          <div>
                            <p className="text-sm font-semibold text-slate-900">{item.title}</p>
                            <p className="mt-1 text-xs leading-5 text-slate-400">{item.detail}</p>
                          </div>
                        </div>
                      </div>
                    ))}
                  </div>
                </div>
              </section>
            ) : (
              <div className="workspace-bottom-reveal pb-4">
                <button
                  onClick={() => setShowBottomPanel(true)}
                  className="inline-flex items-center gap-2 rounded-full border border-[#d7e5ff] bg-white px-4 py-2 text-sm font-semibold text-[#1f6fe5] shadow-[0_12px_24px_rgba(148,163,184,0.12)] transition hover:bg-[#f8fbff]"
                >
                  <ChevronUp className="h-4 w-4" />
                  Pull bottom dock back up
                </button>
              </div>
            )}
          </main>
        </div>

        <footer className="hidden border-t border-sky-100/10 bg-[rgba(15,27,45,0.72)] px-4 py-3 backdrop-blur-xl md:px-6">
          <div className="flex flex-col gap-3 md:flex-row md:items-center md:justify-between">
            <div className="flex flex-wrap items-center gap-4 text-xs text-slate-300">
              <span className="inline-flex items-center gap-2 rounded-full bg-emerald-500/10 px-3 py-1 text-emerald-200">
                <div className="h-2 w-2 rounded-full bg-emerald-400" />
                Copilot ready
              </span>
              <span className="inline-flex items-center gap-2 rounded-full bg-[#121925] px-3 py-1">
                <Rows4 className="h-3.5 w-3.5" />
                {sheet.rowCount} rows x {sheet.columnCount} columns
              </span>
              <span className="inline-flex items-center gap-2 rounded-full bg-[#121925] px-3 py-1">
                <Sigma className="h-3.5 w-3.5" />
                {selectionSummary.label}
              </span>
            </div>
            <div className="flex items-center gap-3 text-xs text-slate-400">
              <span className="inline-flex items-center gap-2">
                <HelpCircle className="h-3.5 w-3.5" />
                Ask Lumina to generate narratives, workflows, and next steps.
              </span>
            </div>
          </div>
        </footer>

        <Suspense fallback={<LazyPanelFallback />}>
          {showCopilot && (
          <div className="absolute inset-0 z-40 md:inset-y-0 md:left-auto md:w-full md:max-w-[28rem]">
            <CopilotSidebar
              onClose={() => setShowCopilot(false)}
              sheet={sheet}
              onUpdateRange={applyRangeValues}
              onApplyFormatting={applyRangeFormatting}
            />
          </div>
        )}

        {showRulesModal && (
          <ConditionalFormattingModal
            rules={sheet.rules}
            onAdd={addRule}
            onRemove={removeRule}
            onClose={() => setShowRulesModal(false)}
          />
        )}

        {showValidationModal && (
          <DataValidationModal
            activeCell={sheet.activeCell}
            currentRule={sheet.activeCell ? sheet.validations[sheet.activeCell] || null : null}
            onSet={(rule) => {
              if (sheet.activeCell) setValidation(sheet.activeCell, rule);
            }}
            onClose={() => setShowValidationModal(false)}
          />
        )}

        {showGenerateModal && (
          <GenerateContentModal
            activeCell={sheet.activeCell}
            onGenerate={handleGenerateContent}
            onClose={() => setShowGenerateModal(false)}
          />
        )}

        {showChartModal && (
          <ChartModal
            data={getChartData()}
            onClose={() => setShowChartModal(false)}
          />
        )}

          {showImportModal && (
            <ImportModal
              onImport={handleImport}
              onClose={() => setShowImportModal(false)}
            />
          )}
        </Suspense>
      </div>
    </div>
  );
};

const MetricCard: React.FC<{ icon: React.ReactNode; eyebrow: string; title: string; detail: string; accent: string }> = ({
  icon,
  eyebrow,
  title,
  detail,
  accent,
}) => (
  <div className={`relative overflow-hidden rounded-[28px] border border-[#e7edf5] bg-white p-4 shadow-[0_16px_34px_rgba(148,163,184,0.12)]`}>
    <div className={`pointer-events-none absolute inset-0 bg-gradient-to-br ${accent}`} />
    <div className="relative flex items-start justify-between gap-3">
      <div>
        <p className="text-[11px] font-semibold uppercase tracking-[0.24em] text-slate-400">{eyebrow}</p>
        <p className="mt-2 text-lg font-bold text-slate-900">{title}</p>
        <p className="mt-1 text-sm text-slate-400">{detail}</p>
      </div>
      <div className="rounded-2xl bg-[#edf4ff] p-2 text-[#1f6fe5]">
        {icon}
      </div>
    </div>
  </div>
);

const HealthRow: React.FC<{ label: string; value: string }> = ({ label, value }) => (
  <div className="flex items-center justify-between rounded-2xl border border-[#edf2f7] bg-white px-3 py-2 text-sm">
    <span className="text-slate-400">{label}</span>
    <span className="font-semibold text-slate-900">{value}</span>
  </div>
);

const FeatureCard: React.FC<{ icon: React.ReactNode; title: string; detail: string }> = ({ icon, title, detail }) => (
  <div className="rounded-3xl border border-[#edf2f7] bg-[#fbfcfe] p-4">
    <div className="mb-3 inline-flex rounded-2xl bg-[#edf4ff] p-2 text-[#1f6fe5]">{icon}</div>
    <h4 className="text-sm font-semibold text-slate-900">{title}</h4>
    <p className="mt-2 text-sm leading-6 text-slate-400">{detail}</p>
  </div>
);

export default App;
