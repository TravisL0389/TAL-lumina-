import { GoogleGenAI, Type } from '@google/genai';

const API_KEY = process.env.GEMINI_API_KEY || process.env.API_KEY || '';
const ai = API_KEY ? new GoogleGenAI({ apiKey: API_KEY }) : null;

const parseCsv = (csvData: string) =>
  csvData
    .trim()
    .split('\n')
    .filter(Boolean)
    .map((row) => row.split(',').map((value) => value.trim()));

const toNumber = (value: string) => {
  const numericValue = Number(value.replace(/[$,%x]/gi, ''));
  return Number.isNaN(numericValue) ? null : numericValue;
};

const localSummary = (csvData: string) => {
  const rows = parseCsv(csvData);
  if (rows.length < 2) return { headers: [], numericColumns: [] as Array<{ header: string; values: number[] }> };

  const headers = rows[0];
  const numericColumns = headers
    .map((header, columnIndex) => {
      const values = rows
        .slice(1)
        .map((row) => toNumber(row[columnIndex] || ''))
        .filter((value): value is number => value !== null);
      return { header, values };
    })
    .filter((column) => column.values.length > 0);

  return { headers, numericColumns };
};

const fallbackQueryAnswer = (query: string, csvData: string) => {
  const normalized = query.toLowerCase();
  const rows = parseCsv(csvData);
  const { headers, numericColumns } = localSummary(csvData);
  const firstNumeric = numericColumns[0];

  if (!rows.length || !headers.length) {
    return 'I need some spreadsheet data before I can answer that. Import or enter a dataset, then ask again.';
  }

  if (normalized.includes('total') || normalized.includes('sum')) {
    if (!firstNumeric) return 'I did not detect a numeric column to total yet.';
    const total = firstNumeric.values.reduce((acc, value) => acc + value, 0);
    return `The total for **${firstNumeric.header}** is **${total.toFixed(2)}**.\n\nYou could express that in-sheet with \`=SUM(${String.fromCharCode(65 + headers.indexOf(firstNumeric.header))}:${String.fromCharCode(65 + headers.indexOf(firstNumeric.header))})\`.`;
  }

  if (normalized.includes('average') || normalized.includes('avg')) {
    if (!firstNumeric) return 'I did not detect a numeric column to average yet.';
    const average = firstNumeric.values.reduce((acc, value) => acc + value, 0) / firstNumeric.values.length;
    return `The average for **${firstNumeric.header}** is **${average.toFixed(2)}**.`;
  }

  if (normalized.includes('summary') || normalized.includes('brief') || normalized.includes('narrative')) {
    const highlights = numericColumns.slice(0, 2).map((column) => {
      const max = Math.max(...column.values);
      const min = Math.min(...column.values);
      return `- **${column.header}** ranges from **${min.toFixed(2)}** to **${max.toFixed(2)}** across the current sheet.`;
    });
    return `### Lumina Brief\n\nThis sheet contains **${rows.length - 1}** live records across **${headers.length}** columns.\n\n${highlights.join('\n') || '- Add more numeric columns to deepen the briefing.'}\n\nThe strongest next step is to run a trend analysis or ask for a role-specific action plan.`;
  }

  return `I can work with this workbook, but there is no live Gemini key configured right now. Try questions like "summarize this sheet", "what is the total", "where are the risks", or "draft an executive brief" and I will give the best local answer available.`;
};

const fallbackAnalysis = (csvData: string) => {
  const rows = parseCsv(csvData);
  const { headers, numericColumns } = localSummary(csvData);
  if (rows.length < 2 || !numericColumns.length) {
    return {
      summary: 'Add or import structured data to unlock trend and anomaly analysis.',
      insights: [],
    };
  }

  const insights = numericColumns.slice(0, 3).map((column, index) => {
    const total = column.values.reduce((acc, value) => acc + value, 0);
    const average = total / column.values.length;
    const max = Math.max(...column.values);
    const min = Math.min(...column.values);
    return {
      title: `${column.header} signal`,
      description:
        index === 0
          ? `Average ${column.header.toLowerCase()} is ${average.toFixed(2)} with a peak of ${max.toFixed(2)} and floor of ${min.toFixed(2)}.`
          : `${column.header} has a spread from ${min.toFixed(2)} to ${max.toFixed(2)} across ${column.values.length} rows.`,
      type: index === 0 ? 'trend' : index === 1 ? 'forecast' : 'risk',
      confidence: 0.52,
      visualization: {
        chartType: 'bar' as const,
        data: column.values.slice(0, 6).map((value, valueIndex) => ({
          label: rows[valueIndex + 1]?.[0] || `Row ${valueIndex + 1}`,
          value,
        })),
        xAxisLabel: headers[0] || 'Row',
        yAxisLabel: column.header,
      },
    };
  });

  return {
    summary: `Detected ${numericColumns.length} numeric measures across ${rows.length - 1} records. The top opportunity is to turn those measures into a scenario comparison or executive brief.`,
    insights,
  };
};

const fillFallbackValues = (prompt: string, targetCells: string[]) => {
  const normalized = prompt.toLowerCase();
  const values: Record<string, string> = {};
  const cityPool = ['Austin', 'Seattle', 'Chicago', 'Atlanta', 'Denver', 'Boston'];

  targetCells.forEach((cellId, index) => {
    if (normalized.includes('date')) {
      values[cellId] = new Date().toISOString().slice(0, 10);
    } else if (normalized.includes('city')) {
      values[cellId] = cityPool[index % cityPool.length];
    } else if (normalized.includes('formula')) {
      values[cellId] = '=SUM(A:A)';
    } else if (targetCells.length === 1) {
      values[cellId] = prompt;
    } else {
      values[cellId] = `${prompt} ${index + 1}`;
    }
  });

  return values;
};

export const copilotChat = async (message: string, context: string, history: { role: string; text: string }[]) => {
  if (!ai) {
    return {
      message: fallbackQueryAnswer(message, context),
      actions: [],
    };
  }

  try {
    const systemInstruction = `You are Lumina, an advanced spreadsheet copilot.
You help users analyze, format, and modify spreadsheets.
Current spreadsheet data (CSV):
${context}

You must respond in JSON using:
{
  "message": "Conversational reply",
  "actions": [
    {
      "type": "format" | "update",
      "range": "A1:B2",
      "style": { "backgroundColor": "#f0f0f0", "fontWeight": "bold", "color": "#111827", "textAlign": "center" },
      "values": [["Value"]]
    }
  ]
}

Use actions only when the user clearly asked for a change.`;

    const contents = history.map((entry) => ({
      role: entry.role === 'user' ? 'user' : 'model',
      parts: [{ text: entry.text }],
    }));
    contents.push({ role: 'user', parts: [{ text: message }] });

    const response = await ai.models.generateContent({
      model: 'gemini-2.5-flash',
      contents,
      config: {
        systemInstruction,
        responseMimeType: 'application/json',
        responseSchema: {
          type: Type.OBJECT,
          properties: {
            message: { type: Type.STRING },
            actions: {
              type: Type.ARRAY,
              items: {
                type: Type.OBJECT,
                properties: {
                  type: { type: Type.STRING, enum: ['format', 'update'] },
                  range: { type: Type.STRING },
                  style: {
                    type: Type.OBJECT,
                    properties: {
                      backgroundColor: { type: Type.STRING },
                      color: { type: Type.STRING },
                      fontWeight: { type: Type.STRING },
                      textAlign: { type: Type.STRING },
                    },
                  },
                  values: {
                    type: Type.ARRAY,
                    items: {
                      type: Type.ARRAY,
                      items: { type: Type.STRING },
                    },
                  },
                },
                required: ['type', 'range'],
              },
            },
          },
          required: ['message'],
        },
      },
    });

    return JSON.parse(response.text || '{"message":"I hit an issue.","actions":[]}');
  } catch (error) {
    console.error('Copilot Error:', error);
    return { message: 'I hit an issue while responding to that request.', actions: [] };
  }
};

export const cleanDataRange = async (csvData: string) => {
  if (!ai) {
    return {
      issues: ['No Gemini key found. Performed a lightweight local cleanup fallback.'],
      cleanedCsv: csvData
        .split('\n')
        .map((row) => row.trim())
        .filter(Boolean)
        .join('\n'),
    };
  }

  try {
    const response = await ai.models.generateContent({
      model: 'gemini-2.5-flash',
      contents: `Review this CSV dataset for duplicates, date inconsistencies, missing values, typos, and formatting problems.
Return JSON with:
- issues: string[]
- cleanedCsv: string

Dataset:
${csvData}`,
      config: {
        responseMimeType: 'application/json',
        responseSchema: {
          type: Type.OBJECT,
          properties: {
            issues: {
              type: Type.ARRAY,
              items: { type: Type.STRING },
            },
            cleanedCsv: { type: Type.STRING },
          },
          required: ['issues', 'cleanedCsv'],
        },
      },
    });

    return JSON.parse(response.text || '{"issues":[],"cleanedCsv":""}');
  } catch (error) {
    console.error('Clean Data Error:', error);
    return { issues: ['Error during cleanup'], cleanedCsv: csvData };
  }
};

export const explainFormula = async (formula: string, context: string) => {
  if (!ai) {
    return `This formula is operating in ${context}. It starts with ${formula.split('(')[0]}, so Lumina would treat it as a spreadsheet calculation and explain the referenced range or logic step by step.`;
  }

  try {
    const response = await ai.models.generateContent({
      model: 'gemini-2.5-flash',
      contents: `Explain this spreadsheet formula in simple terms: "${formula}".
Context: ${context}
Keep it concise and friendly.`,
    });
    return response.text?.trim() || 'Could not explain formula.';
  } catch (error) {
    return 'Error explaining formula.';
  }
};

export const analyzeSheetData = async (csvData: string) => {
  if (!ai) return fallbackAnalysis(csvData);

  try {
    const response = await ai.models.generateContent({
      model: 'gemini-2.5-flash',
      contents: `You are a product-grade spreadsheet intelligence engine.
Analyze this CSV and return:
1. A summary
2. Trends
3. Outliers
4. Correlations
5. Forecasts
6. Risks and opportunities

CSV:
${csvData}`,
      config: {
        responseMimeType: 'application/json',
        responseSchema: {
          type: Type.OBJECT,
          properties: {
            summary: { type: Type.STRING },
            insights: {
              type: Type.ARRAY,
              items: {
                type: Type.OBJECT,
                properties: {
                  title: { type: Type.STRING },
                  description: { type: Type.STRING },
                  type: {
                    type: Type.STRING,
                    enum: ['trend', 'outlier', 'correlation', 'suggestion', 'forecast', 'risk', 'opportunity'],
                  },
                  confidence: { type: Type.NUMBER },
                  visualization: {
                    type: Type.OBJECT,
                    properties: {
                      chartType: { type: Type.STRING, enum: ['bar', 'line', 'pie', 'scatter'] },
                      data: {
                        type: Type.ARRAY,
                        items: {
                          type: Type.OBJECT,
                          properties: {
                            label: { type: Type.STRING },
                            value: { type: Type.NUMBER },
                          },
                        },
                      },
                      xAxisLabel: { type: Type.STRING },
                      yAxisLabel: { type: Type.STRING },
                    },
                  },
                },
                required: ['title', 'description', 'type'],
              },
            },
          },
        },
      },
    });

    return JSON.parse(response.text || '{"summary":"","insights":[]}');
  } catch (error) {
    console.error('Advanced Analysis Error:', error);
    return fallbackAnalysis(csvData);
  }
};

export const answerDataQuery = async (query: string, csvData: string) => {
  if (!ai) return fallbackQueryAnswer(query, csvData);

  try {
    const response = await ai.models.generateContent({
      model: 'gemini-2.5-flash',
      contents: `User query: "${query}"
Spreadsheet data:
${csvData}

Respond directly and use markdown when a table or list helps. If the user wants a summary, make it executive-ready.`,
    });
    return response.text?.trim() || "I couldn't find an answer for that.";
  } catch (error) {
    console.error('Query Error:', error);
    return fallbackQueryAnswer(query, csvData);
  }
};

export const checkFormulaErrors = async (formula: string, context: string) => {
  if (!ai) return null;

  try {
    const response = await ai.models.generateContent({
      model: 'gemini-2.5-flash',
      contents: `Examine this spreadsheet formula: "${formula}".
Context: ${context}
If it is correct, return "CORRECT". Otherwise return only the corrected formula.`,
    });
    const result = response.text?.trim() || '';
    return result === 'CORRECT' ? null : result;
  } catch (error) {
    return null;
  }
};

export const evaluateAiFormula = async (formula: string, context: string) => {
  if (!ai) return '#AI_KEY_REQUIRED';

  try {
    const response = await ai.models.generateContent({
      model: 'gemini-2.5-flash',
      contents: `Evaluate this spreadsheet instruction: "${formula}".
Context: ${context}
Return only the final result.`,
    });
    return response.text?.trim() || 'Error';
  } catch (error) {
    return '#AI_ERROR!';
  }
};

export const suggestFormula = async (request: string, context: string) => {
  if (!ai) {
    if (/average|avg/i.test(request)) return '=AVERAGE(B:B)';
    if (/count/i.test(request)) return '=COUNT(A:A)';
    return '=SUM(B:B)';
  }

  try {
    const response = await ai.models.generateContent({
      model: 'gemini-2.5-flash',
      contents: `You are an expert spreadsheet formula generator.
Request: "${request}"
Context: ${context}
Return only a formula that begins with =.`,
    });
    return (response.text?.trim() || '').replace(/^```=?/, '').replace(/```$/, '').trim();
  } catch (error) {
    return '';
  }
};

export const generateFormattingRule = async (prompt: string) => {
  if (!ai) {
    if (/less|negative|below/i.test(prompt)) {
      return { conditionType: 'lessThan', threshold: '0', style: { backgroundColor: '#3a1616', color: '#fecaca', fontWeight: 'bold' } };
    }
    return { conditionType: 'contains', threshold: 'Watch', style: { backgroundColor: '#332701', color: '#fde68a', fontWeight: 'bold' } };
  }

  try {
    const response = await ai.models.generateContent({
      model: 'gemini-2.5-flash',
      contents: `Generate a conditional formatting rule from: "${prompt}".
Return JSON with conditionType, threshold, and style.`,
      config: {
        responseMimeType: 'application/json',
        responseSchema: {
          type: Type.OBJECT,
          properties: {
            conditionType: { type: Type.STRING, enum: ['greaterThan', 'lessThan', 'equals', 'contains'] },
            threshold: { type: Type.STRING },
            style: {
              type: Type.OBJECT,
              properties: {
                backgroundColor: { type: Type.STRING },
                color: { type: Type.STRING },
                fontWeight: { type: Type.STRING },
              },
            },
          },
          required: ['conditionType', 'threshold', 'style'],
        },
      },
    });
    return JSON.parse(response.text || '{}');
  } catch (error) {
    console.error('Formatting Rule Error:', error);
    return null;
  }
};

export const getFormulaPrediction = async (partialInput: string, context: string) => {
  if (!ai) return '';

  try {
    const response = await ai.models.generateContent({
      model: 'gemini-2.5-flash',
      contents: `Predict the user's intended spreadsheet formula.
Partial input: "${partialInput}"
Context: ${context}
Return only the suggested formula.`,
    });
    return response.text?.trim() || '';
  } catch (error) {
    return '';
  }
};

export const generateCellContent = async (prompt: string, context: string, targetCells: string[]) => {
  if (!ai) return fillFallbackValues(prompt, targetCells);

  try {
    const response = await ai.models.generateContent({
      model: 'gemini-2.5-flash',
      contents: `You are filling spreadsheet cells with AI.
Prompt: "${prompt}"
Context: ${context}
Target cells: ${targetCells.join(', ')}

Return JSON where each key is a target cell id and each value is the exact cell content.`,
      config: {
        responseMimeType: 'application/json',
      },
    });

    return JSON.parse(response.text || '{}');
  } catch (error) {
    return fillFallbackValues(prompt, targetCells);
  }
};

export const flashFillRange = async (prompt: string, context: string) => {
  if (!ai) return {};

  try {
    const response = await ai.models.generateContent({
      model: 'gemini-2.5-flash',
      contents: `Perform a spreadsheet flash fill.
Prompt: "${prompt}"
Context: ${context}
Return JSON where keys are cell ids and values are filled strings.`,
      config: {
        responseMimeType: 'application/json',
      },
    });
    return JSON.parse(response.text || '{}');
  } catch (error) {
    console.error('Flash fill error:', error);
    return {};
  }
};

export const generateValidationRule = async (prompt: string) => {
  if (!ai) {
    const listMatch = prompt.match(/dropdown with (.*)/i);
    const listItems = listMatch ? listMatch[1].split(',').map((item) => item.trim()) : undefined;
    return {
      type: listItems ? 'list' : 'textLength',
      operator: listItems ? undefined : 'greaterThan',
      value1: listItems ? undefined : '0',
      listItems,
      errorMessage: listItems ? 'Choose one of the approved values.' : 'This entry does not match the expected format.',
    };
  }

  try {
    const response = await ai.models.generateContent({
      model: 'gemini-2.5-flash',
      contents: `Translate this spreadsheet validation request into JSON: "${prompt}".
Return:
{
  "type": "type",
  "operator": "operator",
  "value1": "threshold",
  "listItems": ["item1", "item2"],
  "errorMessage": "message"
}`,
      config: {
        responseMimeType: 'application/json',
        responseSchema: {
          type: Type.OBJECT,
          properties: {
            type: { type: Type.STRING },
            operator: { type: Type.STRING },
            value1: { type: Type.STRING },
            listItems: { type: Type.ARRAY, items: { type: Type.STRING } },
            errorMessage: { type: Type.STRING },
          },
          required: ['type', 'errorMessage'],
        },
      },
    });
    return JSON.parse(response.text || '{}');
  } catch (error) {
    return null;
  }
};
