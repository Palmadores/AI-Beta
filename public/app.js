const companyInput = document.getElementById('companyInput');
const resolveBtn = document.getElementById('resolveBtn');
const analyzeBtn = document.getElementById('analyzeBtn');
const downloadBtn = document.getElementById('downloadBtn');
const statusEl = document.getElementById('status');

const progressSection = document.getElementById('progressSection');
const progressTitle = document.getElementById('progressTitle');
const progressList = document.getElementById('progressList');

const ambiguitySection = document.getElementById('ambiguitySection');
const ambiguityList = document.getElementById('ambiguityList');
const resultsSection = document.getElementById('resultsSection');
const resultsTableBody = document.querySelector('#resultsTable tbody');

const confirmAmbiguityBtn = document.getElementById('confirmAmbiguityBtn');
const dialogueSection = document.getElementById('dialogueSection');
const chatCompanySelect = document.getElementById('chatCompanySelect');
const chatMessages = document.getElementById('chatMessages');
const chatInput = document.getElementById('chatInput');
const sendChatBtn = document.getElementById('sendChatBtn');
const voiceBtn = document.getElementById('voiceBtn');

let resolvedCompanies = [];
let ambiguousCompanies = [];
let rows = [];
let progressItems = [];
let chatHistories = {};
let recognition = null;
let voiceActive = false;

function setStatus(message) {
  statusEl.textContent = message;
}

function clamp(value, min, max) {
  const n = Number(value);
  if (Number.isNaN(n)) return min;
  return Math.max(min, Math.min(max, n));
}

function computeAiBeta(row) {
  const disruption = clamp(row.disruption_risk, -1, 0);
  const moat = clamp(row.moat, 0, 1);
  const upside = clamp(row.ai_upside, 0, 1);
  const leverage = clamp(row.ai_leverage, 1, 10);
  return Number(((disruption * (1 - moat)) + (upside * leverage)).toFixed(4));
}

function ragClassForMetric(metric, value) {
  let goodness = 0.5;

  if (metric === 'disruption_risk') {
    goodness = clamp(value + 1, 0, 1);
  } else if (metric === 'moat') {
    goodness = clamp(value, 0, 1);
  } else if (metric === 'ai_upside') {
    goodness = clamp(value, 0, 1);
  } else if (metric === 'ai_leverage') {
    // neutral — high leverage is good for equity, risky for debt
    goodness = 0.5;
  } else if (metric === 'ai_beta') {
    if (value >= 2) return 'rag-good';
    if (value >= 0) return 'rag-mid';
    return 'rag-bad';
  }

  if (goodness >= 0.67) return 'rag-good';
  if (goodness >= 0.34) return 'rag-mid';
  return 'rag-bad';
}

function applyRagClass(element, metric, value) {
  element.classList.remove('rag-good', 'rag-mid', 'rag-bad');
  element.classList.add(ragClassForMetric(metric, Number(value)));
}

function getNamesFromInput() {
  return companyInput.value
    .split('\n')
    .map((s) => s.trim())
    .filter(Boolean);
}

function setProgressMode(title, labels) {
  progressTitle.textContent = title;
  progressItems = labels.map((label, index) => ({
    index,
    label,
    state: 'pending'
  }));
  renderProgress();
}

function progressIcon(item) {
  if (item.state === 'done') return '✓';
  if (item.state === 'in_progress') return '⏳';
  if (item.state === 'review') return '⚠';
  if (item.state === 'error') return '✕';
  return '○';
}

function renderProgress() {
  progressList.innerHTML = '';

  if (!progressItems.length) {
    progressSection.hidden = true;
    return;
  }

  progressSection.hidden = false;

  progressItems.forEach((item) => {
    const li = document.createElement('li');
    li.className = `progress-item ${item.state}`;

    const icon = document.createElement('span');
    icon.className = 'progress-icon';
    icon.textContent = progressIcon(item);

    const label = document.createElement('span');
    label.textContent = item.label;

    li.appendChild(icon);
    li.appendChild(label);
    progressList.appendChild(li);
  });
}

function updateProgress(index, state, label) {
  if (index < 0 || index >= progressItems.length) return;
  progressItems[index].state = state;
  if (label) progressItems[index].label = label;
  renderProgress();
}

function renderAmbiguities() {
  ambiguityList.innerHTML = '';

  if (!ambiguousCompanies.length) {
    ambiguitySection.hidden = true;
    return;
  }

  ambiguitySection.hidden = false;

  ambiguousCompanies.forEach((item, index) => {
    const wrapper = document.createElement('div');
    wrapper.className = 'ambiguity-item';

    const label = document.createElement('label');
    label.textContent = `Input: ${item.input_name}`;
    label.htmlFor = `ambiguous-${index}`;

    const select = document.createElement('select');
    select.id = `ambiguous-${index}`;

    item.candidates.forEach((candidate) => {
      const option = document.createElement('option');
      option.value = JSON.stringify(candidate);
      const tickerPart = candidate.ticker ? ` (${candidate.ticker})` : '';
      const exchangePart = candidate.exchange ? `, ${candidate.exchange}` : '';
      const countryPart = candidate.country ? `, ${candidate.country}` : '';
      option.textContent = `${candidate.name}${tickerPart}${exchangePart}${countryPart} - ${candidate.description || ''}`;
      select.appendChild(option);
    });

    select.addEventListener('change', () => {
      const selected = JSON.parse(select.value);
      item.selected_company = selected;
    });

    if (item.candidates[0]) {
      item.selected_company = item.candidates[0];
    }

    wrapper.appendChild(label);
    wrapper.appendChild(select);
    ambiguityList.appendChild(wrapper);
  });
}

function bindNumberInput(input, rowIndex, key, min, max) {
  input.addEventListener('input', () => {
    rows[rowIndex][key] = clamp(input.value, min, max);
    rows[rowIndex].ai_beta = computeAiBeta(rows[rowIndex]);
    renderResults();
    refreshChatCompanySelect();
  });
}

function renderResults() {
  resultsTableBody.innerHTML = '';
  if (!rows.length) {
    resultsSection.hidden = true;
    dialogueSection.hidden = false;
    chatCompanySelect.innerHTML = '';
    chatCompanySelect.disabled = true;
    chatInput.disabled = true;
    sendChatBtn.disabled = true;
    if (recognition) voiceBtn.disabled = false;
    chatMessages.innerHTML = '<div class="chat-msg assistant"><div class="chat-role">AI</div><div>Run Analyze to populate company rows, then ask questions here about the scoring rationale.</div></div>';
    return;
  }

  resultsSection.hidden = false;
  dialogueSection.hidden = false;
  chatCompanySelect.disabled = false;
  chatInput.disabled = false;
  sendChatBtn.disabled = false;
  if (recognition) voiceBtn.disabled = false;

  rows.forEach((row, index) => {
    const tr = document.createElement('tr');

    const companyTd = document.createElement('td');
    companyTd.textContent = row.company;

    const disruptionTd = document.createElement('td');
    const disruptionInput = document.createElement('input');
    disruptionInput.className = 'grid-input';
    disruptionInput.type = 'number';
    disruptionInput.step = '0.01';
    disruptionInput.min = '-1';
    disruptionInput.max = '0';
    disruptionInput.value = row.disruption_risk;
    bindNumberInput(disruptionInput, index, 'disruption_risk', -1, 0);
    applyRagClass(disruptionInput, 'disruption_risk', row.disruption_risk);
    disruptionTd.appendChild(disruptionInput);

    const moatTd = document.createElement('td');
    const moatInput = document.createElement('input');
    moatInput.className = 'grid-input';
    moatInput.type = 'number';
    moatInput.step = '0.01';
    moatInput.min = '0';
    moatInput.max = '1';
    moatInput.value = row.moat;
    bindNumberInput(moatInput, index, 'moat', 0, 1);
    applyRagClass(moatInput, 'moat', row.moat);
    moatTd.appendChild(moatInput);

    const upsideTd = document.createElement('td');
    const upsideInput = document.createElement('input');
    upsideInput.className = 'grid-input';
    upsideInput.type = 'number';
    upsideInput.step = '0.01';
    upsideInput.min = '0';
    upsideInput.max = '1';
    upsideInput.value = row.ai_upside;
    bindNumberInput(upsideInput, index, 'ai_upside', 0, 1);
    applyRagClass(upsideInput, 'ai_upside', row.ai_upside);
    upsideTd.appendChild(upsideInput);

    const leverageTd = document.createElement('td');
    const leverageInput = document.createElement('input');
    leverageInput.className = 'grid-input';
    leverageInput.type = 'number';
    leverageInput.step = '0.1';
    leverageInput.min = '1';
    leverageInput.max = '10';
    leverageInput.value = row.ai_leverage;
    bindNumberInput(leverageInput, index, 'ai_leverage', 1, 10);
    applyRagClass(leverageInput, 'ai_leverage', row.ai_leverage);
    leverageTd.appendChild(leverageInput);

    const betaTd = document.createElement('td');
    betaTd.className = 'grid-score';
    betaTd.textContent = row.ai_beta;
    applyRagClass(betaTd, 'ai_beta', row.ai_beta);

    const commentTd = document.createElement('td');
    const commentInput = document.createElement('textarea');
    commentInput.className = 'grid-comment';
    commentInput.rows = 3;
    commentInput.value = row.comment || '';
    commentInput.addEventListener('input', () => {
      rows[index].comment = commentInput.value;
      refreshChatCompanySelect();
    });
    commentTd.appendChild(commentInput);

    tr.appendChild(companyTd);
    tr.appendChild(disruptionTd);
    tr.appendChild(moatTd);
    tr.appendChild(upsideTd);
    tr.appendChild(leverageTd);
    tr.appendChild(betaTd);
    tr.appendChild(commentTd);
    resultsTableBody.appendChild(tr);
  });
}

function refreshChatCompanySelect() {
  const current = chatCompanySelect.value;
  chatCompanySelect.innerHTML = '';

  rows.forEach((row, index) => {
    const option = document.createElement('option');
    option.value = String(index);
    option.textContent = row.company;
    chatCompanySelect.appendChild(option);
  });

  if (rows.length) {
    chatCompanySelect.value = current && Number(current) < rows.length ? current : '0';
  }

  renderChatMessages();
}

function chatKeyForSelectedCompany() {
  const idx = Number(chatCompanySelect.value || 0);
  return rows[idx]?.company || '';
}

function addChatMessage(role, text, suggestion) {
  const key = chatKeyForSelectedCompany();
  if (!key) return;
  if (!chatHistories[key]) chatHistories[key] = [];
  chatHistories[key].push({ role, content: text, suggestion: suggestion || null });
  renderChatMessages();
}

function renderChatMessages() {
  const key = chatKeyForSelectedCompany();
  chatMessages.innerHTML = '';

  if (!key) return;

  const history = chatHistories[key] || [];
  history.forEach((entry, i) => {
    const msg = document.createElement('div');
    msg.className = `chat-msg ${entry.role}`;

    const role = document.createElement('div');
    role.className = 'chat-role';
    role.textContent = entry.role === 'assistant' ? 'AI' : 'You';

    const content = document.createElement('div');
    content.textContent = entry.content;

    msg.appendChild(role);
    msg.appendChild(content);

    if (entry.role === 'assistant' && entry.suggestion) {
      const wrap = document.createElement('div');
      wrap.className = 'suggestion-wrap';
      const btn = document.createElement('button');
      btn.type = 'button';
      btn.textContent = 'Apply Suggested Update';
      btn.addEventListener('click', () => {
        applySuggestedUpdate(entry.suggestion);
        entry.suggestion = null;
        renderChatMessages();
      });
      wrap.appendChild(btn);
      msg.appendChild(wrap);
    }

    chatMessages.appendChild(msg);

    if (i === history.length - 1) {
      chatMessages.scrollTop = chatMessages.scrollHeight;
    }
  });
}

function applySuggestedUpdate(suggestion) {
  const index = Number(chatCompanySelect.value || 0);
  if (index < 0 || index >= rows.length) return;

  const row = rows[index];
  const keys = [
    'disruption_risk',
    'moat',
    'ai_upside',
    'ai_leverage',
    'comment'
  ];

  keys.forEach((key) => {
    if (suggestion[key] == null) return;
    row[key] = suggestion[key];
  });

  row.disruption_risk = clamp(row.disruption_risk, -1, 0);
  row.moat = clamp(row.moat, 0, 1);
  row.ai_upside = clamp(row.ai_upside, 0, 1);
  row.ai_leverage = clamp(row.ai_leverage, 1, 10);
  row.ai_beta = computeAiBeta(row);

  renderResults();
  refreshChatCompanySelect();
  setStatus(`Applied suggested update to ${row.company}.`);
}

async function postNdjson(url, body, onEvent) {
  const response = await fetch(url, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(body)
  });

  if (!response.ok) {
    const maybeJson = await response.json().catch(() => ({}));
    throw new Error(maybeJson.error || `Request failed (${response.status})`);
  }

  const reader = response.body.getReader();
  const decoder = new TextDecoder();
  let buffer = '';

  while (true) {
    const { value, done } = await reader.read();
    if (done) break;

    buffer += decoder.decode(value, { stream: true });

    while (true) {
      const newline = buffer.indexOf('\n');
      if (newline < 0) break;

      const line = buffer.slice(0, newline).trim();
      buffer = buffer.slice(newline + 1);

      if (!line) continue;
      onEvent(JSON.parse(line));
    }
  }

  const tail = buffer.trim();
  if (tail) onEvent(JSON.parse(tail));
}

async function postJson(url, body) {
  const response = await fetch(url, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(body)
  });

  const data = await response.json();
  if (!response.ok) {
    throw new Error(data.error || 'Request failed');
  }

  return data;
}

resolveBtn.addEventListener('click', async () => {
  try {
    const names = getNamesFromInput();
    if (!names.length) {
      setStatus('Please add at least one company name.');
      return;
    }

    setStatus('Resolving company names...');
    setProgressMode('Live Progress: Resolve', names);
    resolveBtn.disabled = true;
    analyzeBtn.disabled = true;
    downloadBtn.disabled = true;
    rows = [];
    ambiguousCompanies = [];
    resolvedCompanies = [];
    chatHistories = {};
    renderResults();
    renderAmbiguities();

    let finalResults = [];
    let errors = [];

    await postNdjson('/api/resolve-stream', { names }, (event) => {
      if (event.type === 'started') {
        updateProgress(event.index, 'in_progress', event.input_name || names[event.index]);
      }

      if (event.type === 'progress') {
        if (event.error) {
          updateProgress(event.index, 'error', event.input_name || names[event.index]);
          setStatus(`Resolving ${event.done}/${event.total}: ${event.input_name} (failed)`);
        } else if (event.status === 'ambiguous') {
          updateProgress(event.index, 'review', `${event.input_name} (needs confirmation)`);
          setStatus(`Resolving ${event.done}/${event.total}: ${event.input_name} (ambiguous)`);
        } else {
          updateProgress(event.index, 'done', event.input_name);
          setStatus(`Resolving ${event.done}/${event.total}: ${event.input_name}`);
        }
      }

      if (event.type === 'done') {
        finalResults = event.results || [];
        errors = event.errors || [];
      }

      if (event.type === 'error') {
        throw new Error(event.error || 'Resolve stream failed');
      }
    });

    resolvedCompanies = finalResults
      .filter((r) => r.status === 'resolved' && r.selected_company)
      .map((r) => r.selected_company);

    ambiguousCompanies = finalResults.filter((r) => r.status === 'ambiguous');

    renderAmbiguities();

    const failedPart = errors.length ? ` ${errors.length} failed.` : '';
    if (ambiguousCompanies.length) {
      setStatus(`Resolved ${resolvedCompanies.length}. ${ambiguousCompanies.length} need confirmation.${failedPart}`);
    } else {
      setStatus(`All ${resolvedCompanies.length} company names resolved.${failedPart}`);
    }

    analyzeBtn.disabled = false;
    resolveBtn.disabled = false;

    if (!ambiguousCompanies.length && resolvedCompanies.length) {
      analyzeBtn.click();
    }
  } catch (error) {
    setStatus(`Resolve error: ${error.message}`);
    resolveBtn.disabled = false;
  }
});

analyzeBtn.addEventListener('click', async () => {
  try {
    const confirmedAmbiguous = ambiguousCompanies
      .map((a) => a.selected_company)
      .filter(Boolean);

    const companies = [...resolvedCompanies, ...confirmedAmbiguous];

    if (!companies.length) {
      setStatus('No resolved companies available to analyze.');
      return;
    }

    setStatus(`Analyzing ${companies.length} companies...`);
    setProgressMode('Live Progress: Analyze', companies.map((c) => c.name));
    analyzeBtn.disabled = true;
    resolveBtn.disabled = true;
    downloadBtn.disabled = true;
    rows = [];
    renderResults();

    let finalRows = [];
    let errors = [];

    await postNdjson('/api/analyze-stream', { companies }, (event) => {
      if (event.type === 'started') {
        updateProgress(event.index, 'in_progress', event.company || companies[event.index].name);
      }

      if (event.type === 'progress') {
        if (event.result) {
          rows.push(event.result);
          renderResults();
          refreshChatCompanySelect();
        }

        if (event.error) {
          updateProgress(event.index, 'error', event.company || companies[event.index].name);
          setStatus(`Analyzing ${event.done}/${event.total}: ${event.company} (failed)`);
        } else {
          updateProgress(event.index, 'done', event.company);
          setStatus(`Analyzing ${event.done}/${event.total}: ${event.company}`);
        }
      }

      if (event.type === 'done') {
        finalRows = event.rows || [];
        errors = event.errors || [];
      }

      if (event.type === 'error') {
        throw new Error(event.error || 'Analyze stream failed');
      }
    });

    rows = finalRows.map((r) => ({ ...r, ai_beta: computeAiBeta(r) }));
    renderResults();
    refreshChatCompanySelect();

    const failedPart = errors.length ? ` ${errors.length} failed.` : '';
    setStatus(`Completed analysis for ${rows.length} companies.${failedPart}`);
    downloadBtn.disabled = rows.length === 0;
    analyzeBtn.disabled = false;
    resolveBtn.disabled = false;
  } catch (error) {
    setStatus(`Analysis error: ${error.message}`);
    analyzeBtn.disabled = false;
    resolveBtn.disabled = false;
  }
});

downloadBtn.addEventListener('click', async () => {
  try {
    if (!rows.length) {
      setStatus('No rows to export yet.');
      return;
    }

    setStatus('Preparing Excel download...');

    const response = await fetch('/api/export', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ rows })
    });

    if (!response.ok) {
      const data = await response.json();
      throw new Error(data.error || 'Export failed');
    }

    const blob = await response.blob();
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = 'ai-beta-results.xlsx';
    a.click();
    URL.revokeObjectURL(url);

    setStatus('Excel downloaded.');
  } catch (error) {
    setStatus(`Export error: ${error.message}`);
  }
});

async function sendChat() {
  try {
    const question = chatInput.value.trim();
    if (!question) return;

    const rowIndex = Number(chatCompanySelect.value || 0);
    const row = rows[rowIndex];
    if (!row) {
      setStatus('No company selected for chat.');
      return;
    }

    chatInput.value = '';
    addChatMessage('user', question);
    setStatus(`Asking AI about ${row.company}...`);

    const key = chatKeyForSelectedCompany();
    const history = (chatHistories[key] || []).map((m) => ({ role: m.role, content: m.content }));

    const data = await postJson('/api/dialogue', {
      companyRow: row,
      question,
      history
    });

    addChatMessage('assistant', data.answer || 'No response.', data.suggested_updates || null);
    setStatus(`AI responded for ${row.company}.`);
  } catch (error) {
    setStatus(`Dialogue error: ${error.message}`);
  }
}

confirmAmbiguityBtn.addEventListener('click', () => {
  analyzeBtn.click();
});

sendChatBtn.addEventListener('click', sendChat);
chatInput.addEventListener('keydown', (event) => {
  if (event.key === 'Enter') {
    event.preventDefault();
    sendChat();
  }
});

chatCompanySelect.addEventListener('change', () => {
  renderChatMessages();
});

function setupVoiceInput() {
  const SpeechRecognition = window.SpeechRecognition || window.webkitSpeechRecognition;
  if (!SpeechRecognition) {
    voiceBtn.disabled = true;
    voiceBtn.title = 'Voice not supported in this browser';
    return;
  }

  recognition = new SpeechRecognition();
  recognition.lang = 'en-US';
  recognition.interimResults = true;
  recognition.continuous = false;

  recognition.onstart = () => {
    voiceActive = true;
    voiceBtn.classList.add('voice-active');
    voiceBtn.textContent = 'Listening...';
  };

  recognition.onend = () => {
    voiceActive = false;
    voiceBtn.classList.remove('voice-active');
    voiceBtn.textContent = 'Voice';
  };

  recognition.onresult = (event) => {
    let transcript = '';
    for (let i = event.resultIndex; i < event.results.length; i += 1) {
      transcript += event.results[i][0].transcript;
    }
    chatInput.value = transcript.trim();
  };

  recognition.onerror = () => {
    setStatus('Voice input error. You can still type your question.');
  };

  voiceBtn.addEventListener('click', () => {
    if (!recognition) return;
    if (voiceActive) {
      recognition.stop();
      return;
    }
    recognition.start();
  });
}

setupVoiceInput();
