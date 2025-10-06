const state = {
  sheet: null,
  report: null,
  items: [],
};

const elements = {
  smartsheetToken: document.getElementById('smartsheet-token'),
  smartsheetSheetId: document.getElementById('smartsheet-sheet-id'),
  loadSmartsheet: document.getElementById('load-smartsheet'),
  smartsheetStatus: document.getElementById('smartsheet-status'),
  columnKeyResult: document.getElementById('column-key-result'),
  columnStatus: document.getElementById('column-status'),
  columnDelta: document.getElementById('column-delta'),
  columnBlockers: document.getElementById('column-blockers'),
  columnOwner: document.getElementById('column-owner'),
  meetingNotes: document.getElementById('meeting-notes'),
  googleDocId: document.getElementById('google-doc-id'),
  loadGoogleDoc: document.getElementById('load-google-doc'),
  googleDocStatus: document.getElementById('google-doc-status'),
  toneSelect: document.getElementById('tone-select'),
  customToneWrapper: document.getElementById('custom-tone-wrapper'),
  customToneInput: document.getElementById('custom-tone-input'),
  toneEmphasis: document.getElementById('tone-emphasis'),
  generateReport: document.getElementById('generate-report'),
  reportContent: document.getElementById('report-content'),
  reportWarning: document.getElementById('report-warning'),
  exportPdf: document.getElementById('export-pdf'),
  exportPpt: document.getElementById('export-ppt'),
  exportDocx: document.getElementById('export-docx'),
};

elements.loadSmartsheet.addEventListener('click', async () => {
  const token = elements.smartsheetToken.value.trim();
  const sheetId = elements.smartsheetSheetId.value.trim();

  if (!token || !sheetId) {
    setStatus(elements.smartsheetStatus, 'Provide both access token and sheet ID.', 'error');
    return;
  }

  setStatus(elements.smartsheetStatus, 'Loading sheet…');
  elements.loadSmartsheet.disabled = true;

  try {
    const sheet = await fetchSmartsheetSheet(token, sheetId);
    state.sheet = sheet;
    setStatus(elements.smartsheetStatus, `Loaded \"${sheet.name}\" with ${sheet.totalRowCount} rows.`, 'success');
  } catch (error) {
    console.error(error);
    setStatus(elements.smartsheetStatus, error.message || 'Unable to load Smartsheet data.', 'error');
  } finally {
    elements.loadSmartsheet.disabled = false;
  }
});

elements.loadGoogleDoc.addEventListener('click', async () => {
  const docId = elements.googleDocId.value.trim();
  if (!docId) {
    setStatus(elements.googleDocStatus, 'Provide a Google Doc ID.', 'error');
    return;
  }

  setStatus(elements.googleDocStatus, 'Fetching doc…');
  elements.loadGoogleDoc.disabled = true;

  try {
    const text = await fetchGoogleDoc(docId);
    elements.meetingNotes.value = text;
    setStatus(elements.googleDocStatus, 'Loaded notes from Google Docs.', 'success');
  } catch (error) {
    console.error(error);
    setStatus(elements.googleDocStatus, error.message || 'Unable to load document. Ensure it is shared publicly.', 'error');
  } finally {
    elements.loadGoogleDoc.disabled = false;
  }
});

elements.toneSelect.addEventListener('change', () => {
  const isCustom = elements.toneSelect.value === 'custom';
  elements.customToneWrapper.classList.toggle('hidden', !isCustom);
});

elements.generateReport.addEventListener('click', () => {
  const report = buildReport();
  if (!report) {
    return;
  }

  state.report = report;
  renderReport(report);
  setStatus(elements.reportWarning, '', '');
});

elements.exportPdf.addEventListener('click', () => {
  if (!state.report) {
    flashReportWarning('Generate the report before exporting.');
    return;
  }
  exportPdf(state.report);
});

elements.exportPpt.addEventListener('click', () => {
  if (!state.report) {
    flashReportWarning('Generate the report before exporting.');
    return;
  }
  exportPpt(state.report);
});

elements.exportDocx.addEventListener('click', () => {
  if (!state.report) {
    flashReportWarning('Generate the report before exporting.');
    return;
  }
  exportDocx(state.report);
});

function setStatus(element, message, kind = '') {
  element.textContent = message;
  element.classList.remove('error', 'success');
  if (kind) {
    element.classList.add(kind);
  }
}

async function fetchSmartsheetSheet(token, sheetId) {
  const response = await fetch(`https://api.smartsheet.com/2.0/sheets/${encodeURIComponent(sheetId)}`, {
    headers: {
      Authorization: `Bearer ${token}`,
      Accept: 'application/json',
    },
  });

  if (!response.ok) {
    let message = `Smartsheet request failed (${response.status})`;
    try {
      const data = await response.json();
      if (data && data.message) {
        message = data.message;
      }
    } catch (error) {
      // ignore
    }
    throw new Error(message);
  }

  return response.json();
}

async function fetchGoogleDoc(docId) {
  const response = await fetch(`https://docs.google.com/document/d/${encodeURIComponent(docId)}/export?format=txt`);
  if (!response.ok) {
    throw new Error('Could not fetch document. Ensure link sharing is enabled.');
  }
  return response.text();
}

function buildReport() {
  const notes = elements.meetingNotes.value.trim();
  const tone = elements.toneSelect.value;
  const emphasis = elements.toneEmphasis.value.trim();
  const customTone = elements.customToneInput.value.trim();

  if (!state.sheet) {
    flashReportWarning('Load a Smartsheet sheet before generating the report.');
    return null;
  }

  const mappings = {
    keyResult: elements.columnKeyResult.value.trim(),
    status: elements.columnStatus.value.trim(),
    delta: elements.columnDelta.value.trim(),
    blockers: elements.columnBlockers.value.trim(),
    owner: elements.columnOwner.value.trim(),
  };

  const parsed = parseSmartsheet(state.sheet, mappings);

  const narrative = composeNarrative(parsed, notes, tone, customTone, emphasis);

  const report = {
    generatedAt: new Date(),
    sheetName: state.sheet.name,
    tone: tone === 'custom' ? customTone || 'Custom' : tone,
    emphasis,
    notes,
    items: parsed.items,
    metrics: parsed.metrics,
    blockers: parsed.blockers,
    deltas: parsed.deltas,
    meetingHighlights: extractHighlights(notes),
    narrative,
  };

  return report;
}

function parseSmartsheet(sheet, mappings) {
  const columnsById = new Map(sheet.columns.map((column) => [column.id, column]));
  const columnsByTitle = new Map(sheet.columns.map((column) => [column.title.toLowerCase(), column]));

  function getCellValue(row, columnName) {
    if (!columnName) {
      return '';
    }
    const column = columnsByTitle.get(columnName.toLowerCase());
    if (!column) {
      return '';
    }
    const cell = row.cells.find((c) => c.columnId === column.id);
    if (!cell) {
      return '';
    }
    return cell.displayValue ?? cell.value ?? '';
  }

  const items = [];
  const blockers = [];
  const deltas = [];
  const statusCounter = { green: 0, yellow: 0, red: 0, other: 0 };

  for (const row of sheet.rows) {
    if (row.siblingId) {
      // skip child rows to keep high-level summary concise
      continue;
    }

    const name = getCellValue(row, mappings.keyResult) || getCellValue(row, 'Task Name') || row.cells[0]?.displayValue || 'Untitled';
    const rawStatus = (getCellValue(row, mappings.status) || '').toString().toLowerCase();
    const status = normalizeStatus(rawStatus);
    const delta = getCellValue(row, mappings.delta);
    const blocker = getCellValue(row, mappings.blockers);
    const owner = getCellValue(row, mappings.owner);

    if (blocker) {
      blockers.push({ name, blocker, owner });
    }
    if (delta) {
      deltas.push({ name, delta });
    }

    statusCounter[status] = (statusCounter[status] || 0) + 1;

    items.push({
      name,
      status,
      delta,
      blocker,
      owner,
    });
  }

  const total = items.length || 1;
  const metrics = [
    {
      label: 'Green',
      value: statusCounter.green,
      descriptor: `${Math.round((statusCounter.green / total) * 100)}% on track`,
      tone: 'green',
    },
    {
      label: 'Yellow',
      value: statusCounter.yellow,
      descriptor: `${Math.round((statusCounter.yellow / total) * 100)}% watch list`,
      tone: 'yellow',
    },
    {
      label: 'Red',
      value: statusCounter.red,
      descriptor: `${Math.round((statusCounter.red / total) * 100)}% critical`,
      tone: 'red',
    },
  ];

  return { items, blockers, deltas, metrics };
}

function normalizeStatus(status) {
  if (!status) {
    return 'other';
  }
  if (status.includes('green') || status === 'on track' || status === 'complete') {
    return 'green';
  }
  if (status.includes('yellow') || status.includes('amber') || status.includes('watch')) {
    return 'yellow';
  }
  if (status.includes('red') || status.includes('critical') || status.includes('behind')) {
    return 'red';
  }
  return 'other';
}

function composeNarrative(parsed, notes, tone, customTone, emphasis) {
  const toneConfig = getToneConfig(tone, customTone);
  const totalItems = parsed.items.length;
  const headline = `${parsed.metrics[0].value}/${totalItems} initiatives on track, ${parsed.metrics[2].value} critical`;
  const blockersSummary = parsed.blockers.length
    ? `${parsed.blockers.length} blocker${parsed.blockers.length > 1 ? 's' : ''} need attention`
    : 'No active blockers reported';
  const deltaSummary = parsed.deltas.length
    ? `Key movements: ${parsed.deltas
        .slice(0, 3)
        .map((d) => `${d.name} (${d.delta})`)
        .join('; ')}`
    : 'Minimal movement week-over-week';

  const emphasisText = emphasis ? `Emphasis: ${emphasis}.` : '';

  return `${toneConfig.opening} ${headline}. ${blockersSummary}. ${deltaSummary}. ${emphasisText} ${toneConfig.closing}`.trim();
}

function getToneConfig(tone, customTone) {
  const base = {
    opening: 'Summary:',
    closing: '',
  };

  switch (tone) {
    case 'executive':
      return {
        opening: 'Executive summary:',
        closing: 'Focus decisions on red items and unblock owners before next checkpoint.',
      };
    case 'team':
      return {
        opening: 'Team sync recap:',
        closing: 'Let’s align on owners and next steps to sustain momentum.',
      };
    case 'detailed':
      return {
        opening: 'Detailed update:',
        closing: 'See notes below for supporting context and follow-ups.',
      };
    case 'custom':
      return {
        opening: `${customTone || 'Custom summary'}:`,
        closing: 'Tailored narrative complete.',
      };
    default:
      return base;
  }
}

function extractHighlights(notes) {
  if (!notes) {
    return [];
  }
  const lines = notes
    .split(/\r?\n/)
    .map((line) => line.trim())
    .filter(Boolean);

  const bulletLines = lines.filter((line) => /^[-*•]/.test(line));
  if (bulletLines.length >= 3) {
    return bulletLines.slice(0, 6).map((line) => line.replace(/^[-*•]\s*/, ''));
  }

  return lines.slice(0, 6);
}

function renderReport(report) {
  const container = document.createElement('div');
  container.className = 'report-content';

  container.append(
    renderMetrics(report.metrics),
    renderTrafficLights(report.items),
    renderBlockers(report.blockers),
    renderDeltas(report.deltas),
    renderNarrative(report),
    renderMeetingHighlights(report.meetingHighlights)
  );

  elements.reportContent.replaceChildren(container);
}

function renderMetrics(metrics) {
  const section = document.createElement('section');
  section.className = 'report-section';

  const title = document.createElement('h3');
  title.textContent = 'Overall health';
  section.appendChild(title);

  const grid = document.createElement('div');
  grid.className = 'metrics-grid';

  metrics.forEach((metric) => {
    const card = document.createElement('div');
    card.className = 'metric-card';
    card.innerHTML = `<strong>${metric.value}</strong><span>${metric.label}</span><small>${metric.descriptor}</small>`;
    grid.appendChild(card);
  });

  section.appendChild(grid);
  return section;
}

function renderTrafficLights(items) {
  const section = document.createElement('section');
  section.className = 'report-section';

  const title = document.createElement('h3');
  title.textContent = 'Traffic lights';
  section.appendChild(title);

  const table = document.createElement('table');
  table.className = 'traffic-table';
  const thead = document.createElement('thead');
  thead.innerHTML = '<tr><th>Initiative</th><th>Status</th><th>Delta vs last week</th><th>Owner</th></tr>';
  table.appendChild(thead);

  const tbody = document.createElement('tbody');
  items.slice(0, 12).forEach((item) => {
    const row = document.createElement('tr');
    row.innerHTML = `
      <td>${escapeHtml(item.name)}</td>
      <td>${renderStatusBadge(item.status)}</td>
      <td>${escapeHtml(item.delta || '—')}</td>
      <td>${escapeHtml(item.owner || '—')}</td>
    `;
    tbody.appendChild(row);
  });
  table.appendChild(tbody);
  section.appendChild(table);
  return section;
}

function renderBlockers(blockers) {
  const section = document.createElement('section');
  section.className = 'report-section';
  const title = document.createElement('h3');
  title.textContent = 'Blockers';
  section.appendChild(title);

  if (!blockers.length) {
    const empty = document.createElement('p');
    empty.textContent = 'No blockers reported.';
    section.appendChild(empty);
    return section;
  }

  const list = document.createElement('ul');
  list.className = 'blockers-list';
  blockers.slice(0, 6).forEach((item) => {
    const li = document.createElement('li');
    const owner = item.owner ? ` — Owner: ${escapeHtml(item.owner)}` : '';
    li.innerHTML = `<strong>${escapeHtml(item.name)}</strong>: ${escapeHtml(item.blocker)}${owner}`;
    list.appendChild(li);
  });
  section.appendChild(list);
  return section;
}

function renderDeltas(deltas) {
  const section = document.createElement('section');
  section.className = 'report-section';
  const title = document.createElement('h3');
  title.textContent = 'Delta vs last week';
  section.appendChild(title);

  if (!deltas.length) {
    const empty = document.createElement('p');
    empty.textContent = 'No notable movement captured.';
    section.appendChild(empty);
    return section;
  }

  const list = document.createElement('ul');
  list.className = 'notes-list';
  deltas.slice(0, 6).forEach((delta) => {
    const li = document.createElement('li');
    li.innerHTML = `<strong>${escapeHtml(delta.name)}</strong>: ${escapeHtml(delta.delta)}`;
    list.appendChild(li);
  });
  section.appendChild(list);
  return section;
}

function renderNarrative(report) {
  const section = document.createElement('section');
  section.className = 'report-section';
  const title = document.createElement('h3');
  title.textContent = 'Narrative';
  section.appendChild(title);

  const block = document.createElement('div');
  block.className = 'narrative';
  block.textContent = report.narrative;
  section.appendChild(block);
  return section;
}

function renderMeetingHighlights(highlights) {
  const section = document.createElement('section');
  section.className = 'report-section';
  const title = document.createElement('h3');
  title.textContent = 'Meeting highlights';
  section.appendChild(title);

  if (!highlights.length) {
    const empty = document.createElement('p');
    empty.textContent = 'Add meeting notes to surface highlights.';
    section.appendChild(empty);
    return section;
  }

  const list = document.createElement('ul');
  list.className = 'notes-list';
  highlights.forEach((item) => {
    const li = document.createElement('li');
    li.textContent = item;
    list.appendChild(li);
  });
  section.appendChild(list);
  return section;
}

function renderStatusBadge(status) {
  const classes = {
    green: 'badge green',
    yellow: 'badge yellow',
    red: 'badge red',
    other: 'badge',
  };
  const labels = {
    green: 'Green',
    yellow: 'Yellow',
    red: 'Red',
    other: '—',
  };
  return `<span class="${classes[status]}">● ${labels[status]}</span>`;
}

function escapeHtml(value) {
  const div = document.createElement('div');
  div.textContent = value;
  return div.innerHTML;
}

function flashReportWarning(message) {
  elements.reportWarning.textContent = message;
  elements.reportWarning.classList.remove('hidden');
  elements.reportWarning.classList.add('error');
  elements.reportWarning.scrollIntoView({ behavior: 'smooth', block: 'center' });
  setTimeout(() => elements.reportWarning.classList.add('hidden'), 3500);
}

function exportPdf(report) {
  const doc = new window.jspdf.jsPDF({ unit: 'pt', format: 'a4' });
  const margin = 40;
  let cursorY = margin;
  const lineHeight = 18;

  const pageWidth = doc.internal.pageSize.getWidth() - margin * 2;

  doc.setFont('helvetica', 'bold');
  doc.setFontSize(18);
  doc.text('One Page Report', margin, cursorY);
  cursorY += lineHeight;

  doc.setFont('helvetica', 'normal');
  doc.setFontSize(10);
  doc.text(`Source sheet: ${report.sheetName}`, margin, cursorY);
  cursorY += lineHeight;
  doc.text(`Generated: ${report.generatedAt.toLocaleString()}`, margin, cursorY);
  cursorY += lineHeight * 1.5;

  doc.setFont('helvetica', 'bold');
  doc.setFontSize(12);
  doc.text('Narrative', margin, cursorY);
  cursorY += lineHeight;
  doc.setFont('helvetica', 'normal');
  cursorY = addMultilineText(doc, report.narrative, margin, cursorY, pageWidth, lineHeight, margin);
  cursorY += lineHeight;

  doc.setFont('helvetica', 'bold');
  doc.text('Traffic lights', margin, cursorY);
  cursorY += lineHeight;
  report.items.slice(0, 12).forEach((item) => {
    const line = `${item.name} — ${item.status.toUpperCase()} | Δ ${item.delta || '—'} | Owner ${item.owner || '—'}`;
    cursorY = addMultilineText(doc, line, margin + 10, cursorY, pageWidth - 20, lineHeight, margin);
  });
  cursorY += lineHeight;

  doc.setFont('helvetica', 'bold');
  doc.text('Blockers', margin, cursorY);
  cursorY += lineHeight;
  if (!report.blockers.length) {
    doc.setFont('helvetica', 'normal');
    doc.text('None reported.', margin + 10, cursorY);
    cursorY += lineHeight;
  } else {
    report.blockers.slice(0, 6).forEach((blocker) => {
      const line = `${blocker.name}: ${blocker.blocker}${blocker.owner ? ` (Owner ${blocker.owner})` : ''}`;
      cursorY = addMultilineText(doc, line, margin + 10, cursorY, pageWidth - 20, lineHeight, margin);
    });
  }

  cursorY += lineHeight;
  doc.setFont('helvetica', 'bold');
  doc.text('Meeting highlights', margin, cursorY);
  cursorY += lineHeight;
  if (!report.meetingHighlights.length) {
    doc.setFont('helvetica', 'normal');
    doc.text('Add meeting notes to include highlights.', margin + 10, cursorY);
  } else {
    report.meetingHighlights.forEach((highlight) => {
      cursorY = addMultilineText(doc, `• ${highlight}`, margin + 10, cursorY, pageWidth - 20, lineHeight, margin);
    });
  }

  doc.save('one-page-report.pdf');
}

function addMultilineText(doc, text, x, cursorY, maxWidth, lineHeight, margin) {
  const lines = doc.splitTextToSize(text, maxWidth);
  lines.forEach((line) => {
    if (cursorY > doc.internal.pageSize.getHeight() - margin) {
      doc.addPage();
      cursorY = margin;
    }
    doc.text(line, x, cursorY);
    cursorY += lineHeight;
  });
  return cursorY;
}

function exportPpt(report) {
  const pptx = new window.PptxGenJS();
  const slide = pptx.addSlide();
  slide.addText('One Page Report', { x: 0.5, y: 0.3, fontSize: 24, bold: true });
  slide.addText(`Sheet: ${report.sheetName}`, { x: 0.5, y: 0.9, fontSize: 12 });
  slide.addText(report.narrative, {
    x: 0.5,
    y: 1.2,
    w: 9,
    h: 1.6,
    fontSize: 12,
    color: '2d3748',
    fill: { color: 'f1f5f9' },
    margin: 0.1,
  });

  const tableRows = [
    ['Initiative', 'Status', 'Delta', 'Owner'],
    ...report.items.slice(0, 10).map((item) => [
      truncate(item.name, 60),
      item.status.toUpperCase(),
      truncate(item.delta || '—', 40),
      truncate(item.owner || '—', 20),
    ]),
  ];

  slide.addTable(tableRows, {
    x: 0.5,
    y: 3,
    w: 9,
    colW: [4.2, 1.4, 2.4, 1.2],
    fontSize: 11,
    color: '1f2937',
    fill: 'ffffff',
    valign: 'middle',
    border: { type: 'solid', color: 'd1d5db', pt: 1 },
    bold: true,
    rowH: 0.45,
  });

  slide.addText('Blockers', { x: 0.5, y: 6, fontSize: 14, bold: true });
  slide.addText(
    report.blockers.length
      ? report.blockers
          .slice(0, 4)
          .map((item) => `${truncate(item.name, 30)}: ${truncate(item.blocker, 40)}`)
          .join('\n')
      : 'None reported.',
    { x: 0.5, y: 6.3, fontSize: 11, lineSpacingMultiple: 1.15 }
  );

  pptx.writeFile({ fileName: 'one-page-report.pptx' });
}

function truncate(text, max) {
  if (!text) {
    return text;
  }
  return text.length > max ? `${text.slice(0, max - 1)}…` : text;
}

function exportDocx(report) {
  const { Document, Packer, Paragraph, TextRun } = window.docx;
  const doc = new Document({
    sections: [
      {
        properties: {},
        children: [
          new Paragraph({
            children: [new TextRun({ text: 'One Page Report', bold: true, size: 32 })],
          }),
          new Paragraph({ text: `Sheet: ${report.sheetName}`, spacing: { after: 200 } }),
          new Paragraph({ text: report.narrative, spacing: { after: 300 } }),
          new Paragraph({ text: 'Traffic lights', heading: 'Heading2' }),
          ...report.items.slice(0, 12).map(
            (item) =>
              new Paragraph({
                text: `${item.name} — ${item.status.toUpperCase()} | Δ ${item.delta || '—'} | Owner ${item.owner || '—'}`,
                spacing: { after: 100 },
              })
          ),
          new Paragraph({ text: 'Blockers', heading: 'Heading2' }),
          ...(report.blockers.length
            ? report.blockers.slice(0, 6).map(
                (blocker) =>
                  new Paragraph({
                    text: `${blocker.name}: ${blocker.blocker}${blocker.owner ? ` (Owner ${blocker.owner})` : ''}`,
                    spacing: { after: 100 },
                  })
              )
            : [new Paragraph({ text: 'None reported.', spacing: { after: 100 } })]),
          new Paragraph({ text: 'Meeting highlights', heading: 'Heading2' }),
          ...(report.meetingHighlights.length
            ? report.meetingHighlights.map((highlight) => new Paragraph({ text: `• ${highlight}`, spacing: { after: 100 } }))
            : [new Paragraph({ text: 'Add meeting notes to include highlights.', spacing: { after: 100 } })]),
        ],
      },
    ],
  });

  Packer.toBlob(doc).then((blob) => {
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = 'one-page-report.docx';
    a.click();
    URL.revokeObjectURL(url);
  });
}
