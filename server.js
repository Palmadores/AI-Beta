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
  const timeout = setTimeout(() => controller.abort(), 180000);

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
      console.error(`OpenAI API error ${response.status}:`, body);
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
  const disruption = clamp(scores.disruption_risk, -1, 0);
  const moat = clamp(scores.moat, 0, 1);
  const upside = clamp(scores.ai_upside, 0, 1);
  const leverage = clamp(scores.ai_leverage, 1, 10);
  return Number(((disruption * (1 - moat)) + (upside * leverage)).toFixed(4));
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
        console.error(`Worker error [${index}]:`, error?.message || error);
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
  },
  xai: {
    name: 'xAI Corp',
    ticker: '',
    exchange: 'Private',
    country: 'United States',
    description: 'Elon Musk\'s AI company, developer of the Grok large language model.'
  },
  softbank: {
    name: 'SoftBank Group Corp',
    ticker: '9984',
    exchange: 'TSE',
    country: 'Japan',
    description: 'Japanese tech conglomerate and major AI investor via Vision Fund.'
  },
  coreweave: {
    name: 'CoreWeave, Inc.',
    ticker: '',
    exchange: 'Private',
    country: 'United States',
    description: 'AI cloud infrastructure provider specialising in GPU compute.'
  },
  anthropic: {
    name: 'Anthropic PBC',
    ticker: '',
    exchange: 'Private',
    country: 'United States',
    description: 'AI safety company and developer of the Claude large language model.'
  },
  openai: {
    name: 'OpenAI, Inc.',
    ticker: '',
    exchange: 'Private',
    country: 'United States',
    description: 'AI research company, developer of GPT models and ChatGPT.'
  },
  wpp: {
    name: 'WPP plc',
    ticker: 'WPP',
    exchange: 'LSE',
    country: 'United Kingdom',
    description: 'Global advertising and communications services group.'
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
      'You are a financial/technology analyst scoring AI-Beta for a reinsurance investment team.',
      'Research with web search before scoring.',
      'Return only strict JSON with this shape:',
      '{',
      '  "company": string,',
      '  "scores": {',
      '    "disruption_risk": number,',
      '    "moat": number,',
      '    "ai_upside": number,',
      '    "ai_leverage": number',
      '  },',
      '  "comment": string',
      '}',
      'Scoring bounds and intent:',
      '- disruption_risk: -1 to 0. How severely does AI threaten this company\'s existing revenue and business model?',
      '  -1.0 = AI completely destroys the product/service category (e.g. Chegg displaced by ChatGPT, WPP creative by image AI)',
      '  -0.8 to -0.9 = AI fundamentally disrupts core workflows (e.g. Salesforce CRM replaced by AI agents, Intuit tax by AI, Monday.com by AI project tools)',
      '  -0.4 to -0.6 = meaningful disruption but company can adapt (e.g. traditional banks, legacy software)',
      '  -0.1 to -0.2 = minor disruption, business largely resilient (e.g. physical infrastructure, energy)',
      '  0 = completely impervious (e.g. commodities, sovereign debt)',
      '- moat: 0 to 1. IMPORTANT: this is specifically resistance to AI disruption — NOT general competitive moat.',
      '  Enterprise contracts and brand protect against competitors but NOT against AI replacing the product category.',
      '  0.00-0.05 = AI directly replaces the product, no meaningful defence (Chegg, Duolingo core product, WPP)',
      '  0.05-0.15 = enterprise lock-in or switching costs slow disruption but cannot prevent it (Salesforce, Intuit, Unity)',
      '  0.15-0.35 = some structural advantages slow AI adoption (legacy system integration, compliance requirements)',
      '  0.40-0.65 = genuine regulatory, physical or trust barriers (JPMorgan, insurance, utilities)',
      '  0.70-1.00 = AI fundamentally cannot displace (physical infrastructure, sovereign bonds, critical national infrastructure)',
      '- ai_upside: 0 to 1 (0 = no AI revenue or competitive tailwind; 1 = core AI infrastructure beneficiary e.g. NVIDIA, cloud providers)',
      '- ai_leverage: 0 to 10. Amplification multiplier for structural AI bets only.',
      '  CRITICAL RULE: Adding AI features to an existing product to survive disruption is NOT ai_leverage. That is defensive survival spending.',
      '  ai_leverage is ONLY for companies making structural, debt-funded, existential bets on AI infrastructure winning.',
      '  Ask: "Has this company reorganised its entire capital structure around AI succeeding?" If no → leverage is 0 or near 0.',
      '  0 = being disrupted by AI, or simply unaffected. Includes companies adding AI features defensively (Chegg, Salesforce, Monday.com, WPP, Intuit, Unity)',
      '  0.5-1 = genuine AI product investment but no major capital reorientation (some mid-tier SaaS with real AI R&D)',
      '  2-3 = meaningful AI strategic pivot with real capex commitment (Microsoft, Google — but offset by their scale/diversification)',
      '  4-6 = heavy debt-funded AI capex, business significantly concentrated on AI outcomes (Oracle $40B data center spend, AWS)',
      '  7-10 = pure leveraged AI play, binary outcome — owns the world or bankrupt (xAI, Softbank Vision Fund, CoreWeave)',
      'Formula: AI-Beta = (disruption_risk x (1 - moat)) + (ai_upside x ai_leverage)',
      'Output range: -1 (AI destroys value) to +10 (extreme leveraged AI bet)',
      'CRITICAL RULES:',
      '  1. Score disruption_risk on the LONG-TERM structural threat (3-5 year horizon), not current AI adoption efforts.',
      '     Ask: "Will customers still need this product when AI agents are fully capable?" If no → disruption is -0.8 to -1.0.',
      '  2. A company announcing AI features (copilots, agents, chatbots) to retain customers is being DISRUPTED, not benefiting.',
      '     Salesforce adding Agentforce = evidence of disruption threat, NOT evidence of low disruption risk.',
      '  3. moat = resistance to AI disruption ONLY. Enterprise contracts slow the bleeding but do not prevent structural displacement.',
      '  4. If disruption_risk < -0.5, set ai_leverage = 0 unless the company has made a massive debt-funded infrastructure bet.',
      'Calibration anchors — these are the correct target scores:',
      '  Chegg: disruption=-0.95, moat=0.02, upside=0.02, leverage=0 → AI-Beta ≈ -0.93',
      '  Duolingo: disruption=-0.85, moat=0.05, upside=0.05, leverage=0 → AI-Beta ≈ -0.81',
      '  Salesforce: disruption=-0.87, moat=0.10, upside=0.10, leverage=0 → AI-Beta ≈ -0.78',
      '  Monday.com: disruption=-0.80, moat=0.08, upside=0.08, leverage=0 → AI-Beta ≈ -0.74',
      '  Intuit: disruption=-0.78, moat=0.12, upside=0.10, leverage=0 → AI-Beta ≈ -0.69',
      '  JPMorgan: disruption=-0.30, moat=0.65, upside=0.30, leverage=1 → AI-Beta ≈ +0.20',
      '  Microsoft: disruption=-0.20, moat=0.75, upside=0.75, leverage=2.5 → AI-Beta ≈ +1.83',
      '  NVIDIA: disruption=-0.05, moat=0.85, upside=0.92, leverage=4 → AI-Beta ≈ +3.68',
      '  Oracle: disruption=-0.15, moat=0.65, upside=0.85, leverage=8 → AI-Beta ≈ +6.75',
      '  xAI: disruption=0, moat=0.40, upside=0.95, leverage=9 → AI-Beta ≈ +8.55',
      '  Softbank: disruption=-0.05, moat=0.35, upside=0.90, leverage=10 → AI-Beta ≈ +8.97',
      'Comment rules:',
      '- one sentence only',
      '- explain the key score drivers concisely'
    ].join('\n'),
    user: `Analyze and score all 4 AI-Beta dimensions for: ${descriptor}`
  });

  const scores = output.scores || {};
  const normalizedScores = {
    disruption_risk: clamp(scores.disruption_risk, -1, 0),
    moat: clamp(scores.moat, 0, 1),
    ai_upside: clamp(scores.ai_upside, 0, 1),
    ai_leverage: clamp(scores.ai_leverage, 0, 10)
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
        '    "disruption_risk": number | null,',
        '    "moat": number | null,',
        '    "ai_upside": number | null,',
        '    "ai_leverage": number | null,',
        '    "comment": string | null',
        '  } | null',
        '}',
        'Rules:',
        '- Explain rationale clearly and briefly.',
        '- suggested_updates should be null unless user asks to change/refine score/comment.',
        '- Keep score bounds: disruption_risk [-1,0], moat [0,1] (0=no AI-disruption resistance, 1=fully protected), ai_upside [0,1], ai_leverage [0,10] (0=no AI bet, 10=extreme leveraged AI bet).',
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
        disruption_risk: output.suggested_updates.disruption_risk == null
          ? null
          : clamp(output.suggested_updates.disruption_risk, -1, 0),
        moat: output.suggested_updates.moat == null
          ? null
          : clamp(output.suggested_updates.moat, 0, 1),
        ai_upside: output.suggested_updates.ai_upside == null
          ? null
          : clamp(output.suggested_updates.ai_upside, 0, 1),
        ai_leverage: output.suggested_updates.ai_leverage == null
          ? null
          : clamp(output.suggested_updates.ai_leverage, 0, 10),
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
      'Disruption Risk': r.disruption_risk,
      'Moat': r.moat,
      'AI Upside': r.ai_upside,
      'AI Leverage': r.ai_leverage,
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
