import dotenv from 'dotenv';
import express from 'express';
import path from 'path';
import { fileURLToPath } from 'url';
import XLSX from 'xlsx';

dotenv.config();

const app = express();
const port = process.env.PORT || 3000;
const model = process.env.OPENAI_MODEL || 'gpt-5';
const resolveConcurrency = Math.max(1, Number(process.env.RESOLVE_CONCURRENCY || 3));
const analyzeConcurrency = Math.max(1, Number(process.env.ANALYZE_CONCURRENCY || 2));

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

app.use(express.json({ limit: '2mb' }));
app.use(express.static(path.join(__dirname, 'public')));

function cleanJsonText(text) {
  if (!text || typeof text !== 'string') return '';
  const trimmed = text.trim();
  if (trimmed.startsWith('```')) {
    return trimmed
      .replace(/^```json\s*/i, '')
      .replace(/^```\s*/i, '')
      .replace(/```$/i, '')
      .trim();
  }
  return trimmed;
}

function extractBalancedJsonCandidates(text) {
  const candidates = [];
  const openers = new Set(['{', '[']);

  for (let i = 0; i < text.length; i += 1) {
    if (!openers.has(text[i])) continue;

    const stack = [];
    let inString = false;
    let escaped = false;

    for (let j = i; j < text.length; j += 1) {
      const ch = text[j];

      if (escaped) {
        escaped = false;
        continue;
      }
      if (ch === '\\') {
        escaped = true;
        continue;
      }
      if (ch === '"') {
        inString = !inString;
        continue;
      }
      if (inString) continue;

      if (ch === '{' || ch === '[') stack.push(ch);
      if (ch === '}' || ch === ']') {
        const opener = stack.pop();
        if (!opener) break;
        if ((opener === '{' && ch !== '}') || (opener === '[' && ch !== ']')) break;
        if (stack.length === 0) {
          candidates.push(text.slice(i, j + 1));
          break;
        }
      }
    }
  }

  return candidates;
}

function parseModelJson(raw) {
  if (raw && typeof raw === 'object') return raw;

  const cleaned = cleanJsonText(String(raw || ''));
  const attempts = [cleaned];

  if (cleaned.startsWith('"') && cleaned.endsWith('"')) {
    try {
      const unwrapped = JSON.parse(cleaned);
      if (typeof unwrapped === 'string') attempts.push(unwrapped);
    } catch {
      // ignore
    }
  }

  attempts.push(...extractBalancedJsonCandidates(cleaned));

  for (const attempt of attempts) {
    if (!attempt) continue;
    try {
      return JSON.parse(attempt);
    } catch {
      // try next
    }
  }

  throw new Error('Model did not return valid JSON');
}

function extractResponseText(data) {
  if (typeof data?.output_text === 'string' && data.output_text.trim()) {
    return data.output_text;
  }

  const parts = [];
  for (const out of data?.output || []) {
    if (Array.isArray(out?.content)) {
      for (const c of out.content) {
        if (typeof c?.text === 'string') parts.push(c.text);
        if (typeof c?.output_text === 'string') parts.push(c.output_text);
      }
    }
  }

  return parts.join('\n').trim();
}

async function callOpenAI({ apiKey, system, user, useWebSearch = true }) {
  const controller = new AbortController();
  const timeout = setTimeout(() => controller.abort(), 90000);

  const payload = {
    model,
    input: [
      { role: 'system', content: `${system}\n\nReturn JSON only. No markdown, no prose.` },
      { role: 'user', content: user }
    ]
  };

  if (useWebSearch) {
    payload.tools = [{ type: 'web_search_preview' }];
  }

  try {
    const response = await fetch('https://api.openai.com/v1/responses', {
      method: 'POST',
      headers: {
        Authorization: `Bearer ${apiKey}`,
        'Content-Type': 'application/json'
      },
      body: JSON.stringify(payload),
      signal: controller.signal
    });

    if (!response.ok) {
      const body = await response.text();
      throw new Error(`OpenAI API error ${response.status}: ${body}`);
    }

    return response.json();
  } finally {
    clearTimeout(timeout);
  }
}

async function openAIJson({ system, user, useWebSearch = true }) {
  const apiKey = process.env.OPENAI_API_KEY;
  if (!apiKey) {
    throw new Error('OPENAI_API_KEY is missing. Add it to .env');
  }

  const data = await callOpenAI({ apiKey, system, user, useWebSearch });
  const outputText = extractResponseText(data);

  try {
    return parseModelJson(outputText);
  } catch {
    const repair = await callOpenAI({
      apiKey,
      system: 'You convert content into strict JSON only.',
      user: `Convert this into strict valid JSON only, no markdown:\n\n${outputText || JSON.stringify(data)}`,
      useWebSearch: false
    });

    const repairedText = extractResponseText(repair);
    try {
      return parseModelJson(repairedText);
    } catch {
      const snippet = cleanJsonText(outputText).slice(0, 500);
      throw new Error(`Model did not return valid JSON. Snippet: ${snippet || '(empty)'}`);
    }
  }
}

function clamp(value, min, max) {
  const num = Number(value);
  if (Number.isNaN(num)) return min;
  return Math.max(min, Math.min(max, num));
}

function computeAiBeta(scores) {
  const functional = clamp(scores.functional_susceptibility, -1, 0);
  const digital = clamp(scores.digital_susceptibility, -1, 0);
  const resilience = clamp(scores.resilience, 0, 1);
  const infra = clamp(scores.ai_infrastructure_upside, 0, 1);
  const competitive = clamp(scores.ai_competitiveness_upside, 0, 1);
  return Number((((functional + digital) * resilience) + infra + competitive).toFixed(4));
}

async function mapWithConcurrency(items, concurrency, worker, onProgress, onStart) {
  const results = new Array(items.length);
  let cursor = 0;
  let done = 0;

  async function runWorker() {
    while (true) {
      const index = cursor;
      cursor += 1;
      if (index >= items.length) return;

      if (onStart) {
        onStart({ index, total: items.length, item: items[index] });
      }

      try {
        results[index] = await worker(items[index], index);
      } catch (error) {
        results[index] = { __error: String(error?.message || error) };
      }

      done += 1;
      if (onProgress) {
        onProgress({
          index,
          done,
          total: items.length,
          result: results[index]
        });
      }
    }
  }

  const poolSize = Math.max(1, Math.min(concurrency, items.length));
  const tasks = Array.from({ length: poolSize }, () => runWorker());
  await Promise.all(tasks);

  return results;
}

function initNdjson(res) {
  res.setHeader('Content-Type', 'application/x-ndjson; charset=utf-8');
  res.setHeader('Cache-Control', 'no-cache');
  res.setHeader('Connection', 'keep-alive');
  res.flushHeaders?.();
}

function writeNdjson(res, payload) {
  res.write(`${JSON.stringify(payload)}\n`);
}


const resolveFastPath = {
  apple: {
    name: 'Apple Inc.',
    ticker: 'AAPL',
    exchange: 'NASDAQ',
    country: 'United States',
    description: 'Consumer electronics and software company known for iPhone, Mac, and services.'
  },
  oracle: {
    name: 'Oracle Corporation',
    ticker: 'ORCL',
    exchange: 'NYSE',
    country: 'United States',
    description: 'Enterprise software and cloud infrastructure company.'
  },
  microsoft: {
    name: 'Microsoft Corporation',
    ticker: 'MSFT',
    exchange: 'NASDAQ',
    country: 'United States',
    description: 'Enterprise and consumer software/cloud company.'
  },
  alphabet: {
    name: 'Alphabet Inc.',
    ticker: 'GOOGL',
    exchange: 'NASDAQ',
    country: 'United States',
    description: 'Parent company of Google, focused on internet services and AI.'
  },
  google: {
    name: 'Alphabet Inc.',
    ticker: 'GOOGL',
    exchange: 'NASDAQ',
    country: 'United States',
    description: 'Parent company of Google, focused on internet services and AI.'
  },
  amazon: {
    name: 'Amazon.com, Inc.',
    ticker: 'AMZN',
    exchange: 'NASDAQ',
    country: 'United States',
    description: 'E-commerce and cloud infrastructure company.'
  },
  meta: {
    name: 'Meta Platforms, Inc.',
    ticker: 'META',
    exchange: 'NASDAQ',
    country: 'United States',
    description: 'Social platforms and AI company, parent of Facebook and Instagram.'
  },
  nvidia: {
    name: 'NVIDIA Corporation',
    ticker: 'NVDA',
    exchange: 'NASDAQ',
    country: 'United States',
    description: 'Semiconductor and accelerated computing company.'
  },
  tesla: {
    name: 'Tesla, Inc.',
    ticker: 'TSLA',
    exchange: 'NASDAQ',
    country: 'United States',
    description: 'Electric vehicle and energy technology company.'
  }
};

const resolveCache = new Map();

function normalizeName(name) {
  return String(name || '')
    .toLowerCase()
    .replace(/[.,]/g, ' ')
    .replace(/\b(inc|incorporated|corp|corporation|company|co|ltd|limited|plc|holdings?)\b/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function fromFastPath(name) {
  const key = normalizeName(name);
  const hit = resolveFastPath[key];
  if (!hit) return null;

  return {
    input_name: name,
    status: 'resolved',
    selected_company: { ...hit },
    candidates: [{ id: `${key}-fast`, ...hit }],
    _source: 'fast_path'
  };
}

function normalizeResolvedOutput(name, output, source) {
  return {
    input_name: output.input_name || name,
    status: output.status || 'not_found',
    selected_company: output.selected_company || null,
    candidates: Array.isArray(output.candidates) ? output.candidates : [],
    _source: source
  };
}

async function resolveManyNamesBatch(names) {
  if (!names.length) return [];

  const output = await openAIJson({
    system: [
      'You resolve company names to real public/private business entities.',
      'Use web search to verify and disambiguate.',
      'Return only strict JSON with this shape:',
      '{',
      '  "results": [',
      '    {',
      '      "input_name": string,',
      '      "status": "resolved" | "ambiguous" | "not_found",',
      '      "selected_company": {',
      '        "name": string,',
      '        "ticker": string,',
      '        "exchange": string,',
      '        "country": string,',
      '        "description": string',
      '      } | null,',
      '      "candidates": [',
      '        {',
      '          "id": string,',
      '          "name": string,',
      '          "ticker": string,',
      '          "exchange": string,',
      '          "country": string,',
      '          "description": string',
      '        }',
      '      ]',
      '    }',
      '  ]',
      '}',
      'Rules:',
      '- Keep one output object per input name in same order.',
      '- If ambiguous, provide 2-6 candidates.',
      '- If resolved, selected_company is required and include one candidate.',
      '- Keep descriptions one short sentence.'
    ].join('\n'),
    user: JSON.stringify({ names })
  });

  const mapped = new Map();
  const list = Array.isArray(output.results) ? output.results : [];
  for (const item of list) {
    const key = normalizeName(item?.input_name);
    if (key) mapped.set(key, item);
  }

  return names.map((name) => {
    const item = mapped.get(normalizeName(name));
    if (!item) {
      return {
        input_name: name,
        status: 'not_found',
        selected_company: null,
        candidates: []
      };
    }
    return item;
  });
}

async function resolveNamesWithSpeed(names, hooks = {}) {
  const results = new Array(names.length);
  const unresolved = [];

  for (let i = 0; i < names.length; i += 1) {
    const name = names[i];
    hooks.onStart?.({ index: i, total: names.length, item: name });

    const fast = fromFastPath(name);
    if (fast) {
      results[i] = fast;
      continue;
    }

    const cacheKey = normalizeName(name);
    if (resolveCache.has(cacheKey)) {
      results[i] = { ...resolveCache.get(cacheKey), _source: 'cache' };
      continue;
    }

    unresolved.push({ index: i, name });
  }

  if (unresolved.length) {
    const batchNames = unresolved.map((u) => u.name);
    const batchResolved = await resolveManyNamesBatch(batchNames);

    unresolved.forEach((entry, idx) => {
      const normalized = normalizeResolvedOutput(entry.name, batchResolved[idx] || {}, 'model_batch');
      results[entry.index] = normalized;
      resolveCache.set(normalizeName(entry.name), normalized);
    });
  }

  let done = 0;
  for (let i = 0; i < results.length; i += 1) {
    done += 1;
    hooks.onDoneCount?.({ index: i, done, total: names.length, result: results[i] });
  }

  return results;
}
async function analyzeOneCompany(company) {
  const name = String(company.name || '').trim();

  const descriptor = [
    name,
    company.ticker ? `(ticker: ${company.ticker})` : '',
    company.exchange ? `exchange: ${company.exchange}` : '',
    company.country ? `country: ${company.country}` : ''
  ].filter(Boolean).join(' ');

  const output = await openAIJson({
    system: [
      'You are a financial/technology analyst scoring AI-Beta dimensions through Feb 2028.',
      'Research with web search before scoring.',
      'Return only strict JSON with this shape:',
      '{',
      '  "company": string,',
      '  "scores": {',
      '    "functional_susceptibility": number,',
      '    "digital_susceptibility": number,',
      '    "resilience": number,',
      '    "ai_infrastructure_upside": number,',
      '    "ai_competitiveness_upside": number',
      '  },',
      '  "comment": string',
      '}',
      'Scoring bounds and intent:',
      '- functional_susceptibility: -1 to 0 (-1 = tasks can be quickly replaced by AI, 0 = tasks not currently replaceable by AI)',
      '- digital_susceptibility: -1 to 0 (-1 = very digitally susceptible, 0 = impervious)',
      '- resilience: 0 to 1 (0 = strong moat/high resilience, 1 = no moat/low resilience)',
      '- ai_infrastructure_upside: 0 to 1 (higher = more revenue tied to AI infra)',
      '- ai_competitiveness_upside: 0 to 1 (higher = can gain margin/revenue with AI)',
      'Comment rules:',
      '- one sentence only',
      '- explain your score reasoning clearly with concrete factors'
    ].join('\n'),
    user: `Analyze company and score all 5 dimensions for: ${descriptor}`
  });

  const scores = output.scores || {};
  const normalizedScores = {
    functional_susceptibility: clamp(scores.functional_susceptibility, -1, 0),
    digital_susceptibility: clamp(scores.digital_susceptibility, -1, 0),
    resilience: clamp(scores.resilience, 0, 1),
    ai_infrastructure_upside: clamp(scores.ai_infrastructure_upside, 0, 1),
    ai_competitiveness_upside: clamp(scores.ai_competitiveness_upside, 0, 1)
  };

  return {
    company: output.company || name,
    ...normalizedScores,
    ai_beta: computeAiBeta(normalizedScores),
    comment: String(output.comment || '').trim()
  };
}

app.post('/api/resolve', async (req, res) => {
  try {
    const names = Array.isArray(req.body?.names)
      ? req.body.names.map((n) => String(n || '').trim()).filter(Boolean)
      : [];

    if (!names.length) {
      return res.status(400).json({ error: 'Provide names as a non-empty array.' });
    }

    const results = await resolveNamesWithSpeed(names);
    const errors = results.filter((r) => r?.__error).map((r) => r.__error);

    res.json({
      results: results.filter((r) => !r?.__error),
      errors
    });
  } catch (error) {
    res.status(500).json({ error: error.message });
  }
});

app.post('/api/resolve-stream', async (req, res) => {
  const names = Array.isArray(req.body?.names)
    ? req.body.names.map((n) => String(n || '').trim()).filter(Boolean)
    : [];

  if (!names.length) {
    return res.status(400).json({ error: 'Provide names as a non-empty array.' });
  }

  initNdjson(res);
  writeNdjson(res, { type: 'start', total: names.length });

  try {
    const results = await resolveNamesWithSpeed(names, {
      onStart: ({ index, total, item }) => {
        writeNdjson(res, {
          type: 'started',
          mode: 'resolve',
          index,
          total,
          input_name: item
        });
      },
      onDoneCount: ({ index, done, total, result }) => {
        if (result?.__error) {
          writeNdjson(res, {
            type: 'progress',
            mode: 'resolve',
            index,
            done,
            total,
            input_name: names[index],
            error: result.__error
          });
          return;
        }

        writeNdjson(res, {
          type: 'progress',
          mode: 'resolve',
          index,
          done,
          total,
          input_name: result.input_name,
          status: result.status,
          source: result._source,
          result
        });
      }
    });

    writeNdjson(res, {
      type: 'done',
      mode: 'resolve',
      results: results.filter((r) => !r?.__error),
      errors: results.filter((r) => r?.__error).map((r) => r.__error)
    });
  } catch (error) {
    writeNdjson(res, { type: 'error', mode: 'resolve', error: error.message });
  }

  res.end();
});

app.post('/api/analyze', async (req, res) => {
  try {
    const companies = Array.isArray(req.body?.companies)
      ? req.body.companies.map((c) => ({
          name: String(c?.name || '').trim(),
          ticker: String(c?.ticker || '').trim(),
          exchange: String(c?.exchange || '').trim(),
          country: String(c?.country || '').trim()
        })).filter((c) => c.name)
      : [];

    if (!companies.length) {
      return res.status(400).json({ error: 'Provide companies as a non-empty array.' });
    }

    const rows = await mapWithConcurrency(companies, analyzeConcurrency, analyzeOneCompany);
    const errors = rows.filter((r) => r?.__error).map((r) => r.__error);

    res.json({
      rows: rows.filter((r) => !r?.__error),
      errors
    });
  } catch (error) {
    res.status(500).json({ error: error.message });
  }
});

app.post('/api/analyze-stream', async (req, res) => {
  const companies = Array.isArray(req.body?.companies)
    ? req.body.companies.map((c) => ({
        name: String(c?.name || '').trim(),
        ticker: String(c?.ticker || '').trim(),
        exchange: String(c?.exchange || '').trim(),
        country: String(c?.country || '').trim()
      })).filter((c) => c.name)
    : [];

  if (!companies.length) {
    return res.status(400).json({ error: 'Provide companies as a non-empty array.' });
  }

  initNdjson(res);
  writeNdjson(res, { type: 'start', total: companies.length });

  try {
    const rows = await mapWithConcurrency(companies, analyzeConcurrency, analyzeOneCompany, ({ index, done, total, result }) => {
      if (result?.__error) {
        writeNdjson(res, {
          type: 'progress',
          mode: 'analyze',
          index,
          done,
          total,
          company: companies[index].name,
          error: result.__error
        });
        return;
      }

      writeNdjson(res, {
        type: 'progress',
        mode: 'analyze',
        index,
        done,
        total,
        company: result.company,
        result
      });
    }, ({ index, total, item }) => {
      writeNdjson(res, {
        type: 'started',
        mode: 'analyze',
        index,
        total,
        company: item.name
      });
    });

    writeNdjson(res, {
      type: 'done',
      mode: 'analyze',
      rows: rows.filter((r) => !r?.__error),
      errors: rows.filter((r) => r?.__error).map((r) => r.__error)
    });
  } catch (error) {
    writeNdjson(res, { type: 'error', mode: 'analyze', error: error.message });
  }

  res.end();
});

app.post('/api/dialogue', async (req, res) => {
  try {
    const companyRow = req.body?.companyRow;
    const question = String(req.body?.question || '').trim();
    const history = Array.isArray(req.body?.history) ? req.body.history : [];

    if (!companyRow || !question) {
      return res.status(400).json({ error: 'Provide companyRow and question.' });
    }

    const safeHistory = history.slice(-8).map((m) => ({
      role: m?.role === 'assistant' ? 'assistant' : 'user',
      content: String(m?.content || '').slice(0, 2000)
    }));

    const output = await openAIJson({
      useWebSearch: false,
      system: [
        'You are helping a user review AI-Beta scoring rationale for one company.',
        'Use the provided row context only unless the user asks for new research.',
        'Return strict JSON only with this shape:',
        '{',
        '  "answer": string,',
        '  "suggested_updates": {',
        '    "functional_susceptibility": number | null,',
        '    "digital_susceptibility": number | null,',
        '    "resilience": number | null,',
        '    "ai_infrastructure_upside": number | null,',
        '    "ai_competitiveness_upside": number | null,',
        '    "comment": string | null',
        '  } | null',
        '}',
        'Rules:',
        '- Explain rationale clearly and briefly.',
        '- suggested_updates should be null unless user asks to change/refine score/comment.',
        '- Keep score bounds: functional [-1,0], digital [-1,0], resilience [0,1], infra [0,1], competitiveness [0,1].',
        '- If suggesting updates, only set fields you propose to change; others must be null.'
      ].join('\n'),
      user: JSON.stringify({
        company_row: companyRow,
        conversation_history: safeHistory,
        user_question: question
      })
    });

    let suggested = null;
    if (output?.suggested_updates && typeof output.suggested_updates === 'object') {
      suggested = {
        functional_susceptibility: output.suggested_updates.functional_susceptibility == null
          ? null
          : clamp(output.suggested_updates.functional_susceptibility, -1, 0),
        digital_susceptibility: output.suggested_updates.digital_susceptibility == null
          ? null
          : clamp(output.suggested_updates.digital_susceptibility, -1, 0),
        resilience: output.suggested_updates.resilience == null
          ? null
          : clamp(output.suggested_updates.resilience, 0, 1),
        ai_infrastructure_upside: output.suggested_updates.ai_infrastructure_upside == null
          ? null
          : clamp(output.suggested_updates.ai_infrastructure_upside, 0, 1),
        ai_competitiveness_upside: output.suggested_updates.ai_competitiveness_upside == null
          ? null
          : clamp(output.suggested_updates.ai_competitiveness_upside, 0, 1),
        comment: output.suggested_updates.comment == null
          ? null
          : String(output.suggested_updates.comment).trim()
      };
    }

    res.json({
      answer: String(output?.answer || '').trim(),
      suggested_updates: suggested
    });
  } catch (error) {
    res.status(500).json({ error: error.message });
  }
});

app.post('/api/export', (req, res) => {
  try {
    const { rows } = req.body;
    if (!Array.isArray(rows) || rows.length === 0) {
      return res.status(400).json({ error: 'Provide rows as a non-empty array.' });
    }

    const worksheetRows = rows.map((r) => ({
      Company: r.company,
      'Functional Susceptibility': r.functional_susceptibility,
      'Digital Susceptibility': r.digital_susceptibility,
      Resilience: r.resilience,
      'AI Infrastructure Upside': r.ai_infrastructure_upside,
      'AI Competitiveness Upside': r.ai_competitiveness_upside,
      'AI-Beta Score': r.ai_beta,
      Comment: r.comment
    }));

    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.json_to_sheet(worksheetRows);
    XLSX.utils.book_append_sheet(wb, ws, 'AI-Beta Results');

    const buffer = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });

    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', 'attachment; filename="ai-beta-results.xlsx"');
    res.send(buffer);
  } catch (error) {
    res.status(500).json({ error: error.message });
  }
});

app.get('*', (_req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'index.html'));
});

app.listen(port, () => {
  console.log(`AI-Beta app running on http://localhost:${port}`);
});
