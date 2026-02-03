const REQUIRED_HEADERS = [
  "Reviewer",
  "File",
  "Why Law?",
  "Thrive?",
  "Contribute?",
  "Know?",
  "Tags",
  "Reach",
  "Target",
  "Safety",
  "Notes",
  "Anything Else?",
];

const RUBRIC_TO_SCORE = {
  "Strongly Disagree": 1,
  Disagree: 2,
  Neutral: 3,
  Agree: 4,
  "Strongly Agree": 5,
};

const BAND_LABELS = {
  T14: "Top 14",
  T25: "Top 25",
  T50: "Top 50",
  T100: "Top 100",
  "T100+": "Top 100+",
};

const TAG_DEFINITIONS = [
  { name: "Academics Need Explanation", polarity: "negative" },
  { name: "Repetitive Statements", polarity: "negative" },
  { name: "Generic Why X", polarity: "negative" },
  { name: "Writing Needs Improvement", polarity: "negative" },
  { name: "Too Much Early Life", polarity: "negative" },
  { name: "Impressive Experiences", polarity: "positive" },
  { name: "Lovely PS", polarity: "positive" },
  { name: "Likable", polarity: "positive" },
  { name: "Overcame Challenges", polarity: "positive" },
  { name: "Unique Perspective", polarity: "positive" },
  { name: "Good Writer", polarity: "positive" },
  { name: "Useful addendum", polarity: "positive" },
  { name: "Be More Professional", polarity: "negative" },
  { name: "Moving Story", polarity: "positive" },
  { name: "Good Community Member", polarity: "positive" },
  { name: "Test History Needs Explanation", polarity: "negative" },
  { name: "Workplace-Bound", polarity: "negative" },
  { name: "Questionable Formatting", polarity: "negative" },
  { name: "Lovely DS", polarity: "positive" },
  { name: "Unnecessary Addendum", polarity: "negative" },
  { name: "Lovely Supplemental", polarity: "positive" },
  { name: "Rigorous Undergrad", polarity: "positive" },
  { name: "Leadership Evidence", polarity: "positive" },
  { name: "Strong Service Record", polarity: "positive" },
  { name: "Clear Career Vision", polarity: "positive" },
  { name: "Cohesive Narrative", polarity: "positive" },
  { name: "Strong Recommendation Signals", polarity: "positive" },
  { name: "Compelling Why Law", polarity: "positive" },
  { name: "Thin Why Law Narrative", polarity: "negative" },
  { name: "Limited Impact Evidence", polarity: "negative" },
  { name: "Overly Broad Goals", polarity: "negative" },
  { name: "Resume-PS Mismatch", polarity: "negative" },
  { name: "Needs More Specificity", polarity: "negative" },
  { name: "Outstanding Writing Voice", polarity: "positive" },
  { name: "Strong Professional Maturity", polarity: "positive" },
  { name: "Interview Ready", polarity: "positive" },
  { name: "Application Timing Risk", polarity: "negative" },
  { name: "Score Band Risk", polarity: "negative" },
  { name: "GPA Context Needed", polarity: "negative" },
  { name: "Exceptional Fit Signals", polarity: "positive" },
];

const TAG_NAME_SET = new Set(TAG_DEFINITIONS.map((tag) => tag.name));

const state = {
  fileName: "",
  parsedRows: [],
  groupedRows: new Map(),
  availableStudents: [],
  studentInputByFile: {},
  reports: [],
  currentReport: null,
  warnings: [],
  errors: [],
};

let fitResizeScheduled = false;

const PRINT_CSS = `
  @import url("https://fonts.googleapis.com/css2?family=Fraunces:wght@600;700&family=Lexend:wght@400;500;600;700;800&display=swap");
  @page { size: 8.5in 11in; margin: 0; }
  body { margin: 0; font-family: "Lexend", "Segoe UI", Tahoma, sans-serif; color: #202530; }
  .doc-shell { background: #fff; }
  .page { width: auto; min-height: 10in; padding: 0.5in; margin: 0; break-after: page; }
  .page.summary-page { padding: 0; min-height: 792px; height: 792px; width: 612px; transform: scale(1.3333333333); transform-origin: top left; margin-bottom: 264px; }
  .page:last-child { break-after: auto; }
  .page-header { border-bottom: 2px solid #111; padding-bottom: 0.1in; margin-bottom: 0.2in; }
  .doc-title { margin: 0; font-size: 20pt; }
  .subtitle { margin: 4px 0 0; color: #5e6778; font-size: 11pt; }
  .row { display: grid; grid-template-columns: 180px 1fr; gap: 10px; margin-bottom: 10px; align-items: baseline; }
  .label { font-weight: 600; }
  .value { min-height: 1.2em; }
  .small { font-size: 10pt; color: #5e6778; }
  .rating-row { border: 1px solid #d8dee9; border-radius: 8px; padding: 8px 10px; margin-bottom: 8px; }
  .stars-wrap { display: flex; align-items: center; gap: 8px; margin-top: 6px; }
  .stars { display: inline-flex; gap: 4px; }
  .star { width: 18px; height: 18px; }
  .star.filled polygon { fill: #f4b400; stroke: #b78600; }
  .star.empty polygon { fill: #eef2f7; stroke: #a6b1c2; }
  .rating-number { font-weight: 600; font-size: 12px; }
  .band-list { margin: 0; padding-left: 18px; }
  .band-list li { margin-bottom: 6px; }
  .tag-grid { display: grid; grid-template-columns: repeat(3, minmax(0, 1fr)); gap: 8px; }
  .tag-pill { border: 1px solid #c4cddd; border-radius: 999px; padding: 6px 10px; font-size: 10px; line-height: 1.1; text-align: center; display: flex; align-items: center; justify-content: center; position: relative; overflow: visible; }
  .tag-badges { display: inline-flex; gap: 4px; position: absolute; top: -8px; right: 8px; margin-left: 0; vertical-align: baseline; }
  .tag-badge { min-width: 16px; height: 16px; border-radius: 999px; display: inline-flex; align-items: center; justify-content: center; font-size: 10px; font-weight: 700; background: #111827; color: #fff; border: 1px solid #fff; }
  .tag-pill.inactive { color: #5e6778; background: #fff; }
  .tag-pill.active-positive { background: #dcfce7; border-color: #22c55e; color: #14532d; }
  .tag-pill.active-negative { background: #fee2e2; border-color: #ef4444; color: #7f1d1d; }
  .reader-card { border: 1px solid #d8dee9; border-radius: 8px; padding: 10px; margin-bottom: 10px; }
  .reader-title { margin: 0 0 8px; font-size: 14px; }
  .notes-box { border: 1px solid #333; min-height: 1.5in; padding: 10px; white-space: pre-wrap; margin-bottom: 8px; }
  .hero-head { display: flex; justify-content: space-between; align-items: center; height: 56px; background: #f7f5f3; padding: 0 18px; margin-bottom: 0; border-radius: 6px 6px 0 0; }
  .brand-left { font-size: 18px; font-weight: 800; color: #0f172a; }
  .brand-left .badge { background: #f3ce73; border-radius: 6px; padding: 2px 8px; margin-right: 6px; }
  .brand-right { font-family: "Fraunces", "Times New Roman", serif; color: #14b8a6; font-size: 18px; font-weight: 700; }
  .summary-banner { height: 50px; background: linear-gradient(-45deg, #15b79e 0%, #227f9c 100%); color: #fcfaf8; text-align: center; font-family: "Fraunces", "Times New Roman", serif; font-size: 28px; font-weight: 700; line-height: 50px; margin-bottom: 0; }
  .section-block { display: grid; grid-template-columns: 40px 1fr; gap: 0; margin-bottom: 0; width: 100%; background: #fff; border-top: 1px solid #d8dee9; border-bottom: 1px solid #d8dee9; }
  .section-block.readers-section { height: 185px; }
  .section-block.readers-section .section-body { height: 100%; overflow: hidden; }
  .section-block.readers-section .rail-label { height: calc(100% - 40px); margin: 20px 0; display: flex; align-items: center; justify-content: center; padding: 0; overflow: hidden; font-size: 17px; }
  .section-block.readers-section .avatar { width: 88px; height: 88px; border-radius: 20px; font-size: 24px; }
  .section-block.readers-section .reader-name { font-size: 28px; }
  .section-block.readers-section .reader-bio { font-size: 16px; }
  .section-block.readers-section .section-body { padding: 12px 16px; }
  .section-block.key-info-section { height: 212px; }
  .section-block.key-info-section .section-body { height: 100%; }
  .section-block.key-info-section .rail-label { height: calc(100% - 40px); margin: 20px 0; display: flex; align-items: center; justify-content: center; padding: 0; overflow: hidden; font-size: 17px; }
  .section-block.key-info-section .section-body { padding: 14px 16px; }
  .section-block.key-info-section .key-top-item { font-size: 14px; padding: 8px 10px; }
  .section-block.key-info-section .metric-col { padding: 6px 8px; }
  .section-block.key-info-section .metric-title { font-size: 14px; margin: 0 0 4px; }
  .section-block.key-info-section .small { font-size: 9px; }
  .section-block.key-info-section .compact-stars .star { width: 12px; height: 12px; }
  .section-block.key-info-section .band-row { font-size: 9px; gap: 4px; padding: 3px 6px; grid-template-columns: 56px repeat(5, minmax(0, 1fr)); }
  .section-block.tags-section { height: 289px; }
  .section-block.tags-section .section-body { height: 100%; }
  .section-block.tags-section .rail-label { height: calc(100% - 40px); margin: 20px 0; display: flex; align-items: center; justify-content: center; padding: 0; overflow: hidden; font-size: 17px; }
  .section-block.tags-section .section-body { padding: 14px 16px; }
  .rail-label { writing-mode: vertical-rl; transform: rotate(180deg); background: #334155; color: #fff; width: 40px; border-radius: 18px 0 0 18px; text-align: center; font-family: "Fraunces", "Times New Roman", serif; font-size: 22px; font-weight: 700; letter-spacing: 0.4px; padding: 16px 8px; }
  .section-body { border: 0; border-radius: 0; padding: 14px 16px; background: #fff; width: 100%; min-width: 0; overflow: hidden; }
  .fit-content { transform-origin: top left; width: 100%; }
  .reader-grid { display: grid; grid-template-columns: repeat(3, minmax(0, 1fr)); gap: 12px; }
  .reader-col { padding: 4px 8px; border-right: 1px solid #d8dee9; }
  .reader-col:last-child { border-right: 0; }
  .avatar { width: 72px; height: 72px; border-radius: 16px; margin: 0 auto 6px; display: flex; align-items: center; justify-content: center; background: linear-gradient(135deg, #2b7abf, #14b8a6); color: #fff; font-weight: 800; font-size: 20px; }
  .reader-name { text-align: center; margin: 0 0 6px; font-size: 22px; font-family: "Fraunces", "Times New Roman", serif; }
  .reader-bio { margin: 0; text-align: center; font-size: 11px; line-height: 1.35; }
  .key-card { border: 1px solid #d8dee9; border-radius: 14px; overflow: hidden; width: 100%; }
  .key-top { display: grid; grid-template-columns: repeat(3, minmax(0, 1fr)); border-bottom: 1px solid #d8dee9; }
  .key-top-item { padding: 8px 12px; font-size: 19px; font-weight: 700; border-right: 1px solid #d8dee9; }
  .key-top-label { opacity: 0.9; }
  .key-top-value { margin-left: 4px; font-weight: 800; }
  .key-top-item:last-child { border-right: 0; }
  .metric-grid { display: grid; grid-template-columns: repeat(4, minmax(0, 1fr)); border-bottom: 1px solid #d8dee9; }
  .metric-col { padding: 8px 10px; border-right: 1px solid #d8dee9; text-align: center; }
  .metric-col:last-child { border-right: 0; }
  .metric-title { margin: 0 0 6px; font-size: 20px; font-weight: 700; }
  .compact-stars { display: inline-flex; gap: 3px; justify-content: center; }
  .compact-stars .star { width: 14px; height: 14px; }
  .band-row { display: grid; grid-template-columns: 62px repeat(5, minmax(0, 1fr)); align-items: center; gap: 6px; padding: 5px 8px; border-top: 1px solid #e5e7eb; font-size: 11px; }
  .band-row:first-child { border-top: 0; }
  .band-name { font-weight: 700; border-radius: 999px; text-align: center; padding: 2px 6px; }
  .band-name.reach { background: #fff2df; color: #b45309; }
  .band-name.target { background: #e0f2fe; color: #0c4a6e; }
  .band-name.safety { background: #dcfce7; color: #14532d; }
  .band-chip { border: 1px solid #d1d5db; border-radius: 8px; text-align: center; padding: 2px 0; font-weight: 700; }
  .band-chip.active { background: #dbeafe; border-color: #60a5fa; }
  .design-tag-grid { display: grid; grid-template-columns: repeat(4, minmax(0, 1fr)); gap: 8px 10px; }
`;

const PRINT_FIT_SCRIPT = `
  (() => {
    function fitSectionContent(section) {
      const container = section.querySelector(".section-body");
      const content = container?.querySelector(".fit-content");
      if (!container || !content) return;

      content.style.transform = "scale(1)";
      content.style.width = "100%";

      const availableHeight = container.clientHeight;
      const availableWidth = container.clientWidth;
      const neededHeight = content.scrollHeight;
      const neededWidth = content.scrollWidth;

      if (!availableHeight || !availableWidth || !neededHeight || !neededWidth) return;

      const heightScale = availableHeight / neededHeight;
      const widthScale = availableWidth / neededWidth;
      const scale = Math.min(1, heightScale, widthScale);

      if (scale < 1) {
        content.style.transform = "scale(" + scale + ")";
        content.style.width = (100 / scale) + "%";
      }
    }

    function fitAllSummarySections() {
      document.querySelectorAll(".summary-page .section-block").forEach(fitSectionContent);
    }

    window.addEventListener("load", () => {
      fitAllSummarySections();
      setTimeout(() => {
        fitAllSummarySections();
        window.focus();
        window.print();
      }, 80);
    });
  })();
`;

const csvFileInput = document.getElementById("csvFile");
const generateBtn = document.getElementById("generateBtn");
const studentSelect = document.getElementById("studentSelect");
const lsatInput = document.getElementById("lsatInput");
const gpaInput = document.getElementById("gpaInput");
const kjdSelect = document.getElementById("kjdSelect");
const urmSelect = document.getElementById("urmSelect");
const inputHintEl = document.getElementById("inputHint");
const downloadPdfBtn = document.getElementById("downloadPdfBtn");
const printCurrentBtn = document.getElementById("printCurrentBtn");
const statusEl = document.getElementById("status");
const validationEl = document.getElementById("validation");
const previewRoot = document.getElementById("previewRoot");

csvFileInput.addEventListener("change", onCsvSelected);
generateBtn.addEventListener("click", onGenerateDocuments);
studentSelect.addEventListener("change", onStudentSelectionChange);
lsatInput.addEventListener("input", onManualInputChange);
gpaInput.addEventListener("input", onManualInputChange);
kjdSelect.addEventListener("change", onManualInputChange);
urmSelect.addEventListener("change", onManualInputChange);
downloadPdfBtn.addEventListener("click", onDownloadCurrentStudentPdf);
printCurrentBtn.addEventListener("click", onPrintCurrentStudent);
window.addEventListener("resize", onWindowResize);

function setStatus(message) {
  statusEl.textContent = message;
}

function escapeHtml(value) {
  return String(value ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#39;");
}

function showValidationMessages() {
  const blocks = [];
  if (state.errors.length) {
    blocks.push(
      `<div class="error"><strong>Errors:</strong> ${escapeHtml(
        state.errors.join(" | ")
      )}</div>`
    );
  }
  if (state.warnings.length) {
    blocks.push(
      `<div class="warn"><strong>Warnings:</strong> ${escapeHtml(
        state.warnings.join(" | ")
      )}</div>`
    );
  }
  validationEl.innerHTML = blocks.length ? blocks.join("") : "No validation issues.";
}

function clearGeneratedResults({ resetSelector } = { resetSelector: false }) {
  state.reports = [];
  state.currentReport = null;
  state.warnings = [];
  state.errors = [];
  previewRoot.innerHTML =
    '<p class="placeholder">Upload a CSV and generate documents to preview output.</p>';
  if (resetSelector) {
    studentSelect.disabled = true;
    studentSelect.innerHTML = "<option>Upload CSV first</option>";
  }
  downloadPdfBtn.disabled = true;
  printCurrentBtn.disabled = true;
}

function parseCsv(csvText) {
  const rows = [];
  let row = [];
  let cell = "";
  let inQuotes = false;

  for (let i = 0; i < csvText.length; i += 1) {
    const ch = csvText[i];
    const next = csvText[i + 1];

    if (ch === '"') {
      if (inQuotes && next === '"') {
        cell += '"';
        i += 1;
      } else {
        inQuotes = !inQuotes;
      }
      continue;
    }

    if (!inQuotes && ch === ",") {
      row.push(cell);
      cell = "";
      continue;
    }

    if (!inQuotes && (ch === "\n" || ch === "\r")) {
      if (ch === "\r" && next === "\n") i += 1;
      row.push(cell);
      rows.push(row);
      row = [];
      cell = "";
      continue;
    }

    cell += ch;
  }

  if (cell.length || row.length) {
    row.push(cell);
    rows.push(row);
  }

  return rows.filter((r) => r.some((c) => c.trim().length > 0));
}

function readCsvRows(csvText) {
  const allRows = parseCsv(csvText);
  if (!allRows.length) throw new Error("CSV is empty.");

  const headers = allRows[0].map((h) => h.trim());
  if (!headers.length || headers.every((h) => h.length === 0)) {
    throw new Error("CSV is missing a header row.");
  }

  const rows = allRows.slice(1).map((rowValues) => {
    const row = {};
    headers.forEach((header, index) => {
      row[header] = (rowValues[index] ?? "").trim();
    });
    return row;
  });

  return { headers, rows };
}

function validateHeaders(headers) {
  return REQUIRED_HEADERS.filter((header) => !headers.includes(header));
}

function groupRowsByFile(rows) {
  const groups = new Map();
  let blankFileRows = 0;

  rows.forEach((row) => {
    const file = (row.File || "").trim();
    if (!file) {
      blankFileRows += 1;
      return;
    }
    if (!groups.has(file)) groups.set(file, []);
    groups.get(file).push(row);
  });

  return { groups, blankFileRows };
}

function shuffleReaders(rows) {
  const shuffled = [...rows];
  for (let i = shuffled.length - 1; i > 0; i -= 1) {
    const j = Math.floor(Math.random() * (i + 1));
    [shuffled[i], shuffled[j]] = [shuffled[j], shuffled[i]];
  }
  return shuffled;
}

function mapRubricToScore(value, fileName, fieldLabel) {
  const trimmed = String(value || "").trim();
  if (!trimmed) return null;
  const score = RUBRIC_TO_SCORE[trimmed];
  if (!score) {
    state.warnings.push(`${fileName}: unknown ${fieldLabel} value "${trimmed}"`);
    return null;
  }
  return score;
}

function averageScores(scores) {
  if (!scores.length) return null;
  return scores.reduce((sum, value) => sum + value, 0) / scores.length;
}

function parseTags(csvTagString) {
  return String(csvTagString || "")
    .split(",")
    .map((tag) => tag.trim())
    .filter(Boolean);
}

function mapBandValue(value, fileName, fieldLabel) {
  const trimmed = String(value || "").trim();
  if (!trimmed) return "";
  if (BAND_LABELS[trimmed]) return trimmed;
  state.warnings.push(`${fileName}: unknown ${fieldLabel} value "${trimmed}"`);
  return `Unknown band: ${trimmed}`;
}

function uniqueNonEmpty(values) {
  return [...new Set(values.map((v) => String(v || "").trim()).filter(Boolean))];
}

function buildStudentReport(fileName, rows, manual) {
  const randomizedRows = shuffleReaders(rows);
  const labeledReaders = randomizedRows.map((row, index) => ({
    row,
    label: `Reader ${index + 1}`,
  }));
  const readers = labeledReaders.map(({ row, label }) => ({
    label,
    notes: row.Notes || "",
    anythingElse: row["Anything Else?"] || "",
    rawReviewer: row.Reviewer || "",
  }));

  const whyLawScores = rows
    .map((row) => mapRubricToScore(row["Why Law?"], fileName, "Why Law?"))
    .filter((score) => score !== null);
  const thriveScores = rows
    .map((row) => mapRubricToScore(row["Thrive?"], fileName, "Thrive?"))
    .filter((score) => score !== null);
  const contributeScores = rows
    .map((row) => mapRubricToScore(row["Contribute?"], fileName, "Contribute?"))
    .filter((score) => score !== null);
  const knowScores = rows
    .map((row) => mapRubricToScore(row["Know?"], fileName, "Know?"))
    .filter((score) => score !== null);

  const unknownTags = new Set();
  const activeTags = new Set();
  const tagReaderMap = new Map();
  labeledReaders.forEach(({ row, label }) => {
    parseTags(row.Tags).forEach((tag) => {
      if (TAG_NAME_SET.has(tag)) {
        activeTags.add(tag);
        if (!tagReaderMap.has(tag)) tagReaderMap.set(tag, new Set());
        tagReaderMap.get(tag).add(label);
      } else {
        unknownTags.add(tag);
      }
    });
  });

  if (unknownTags.size) {
    state.warnings.push(
      `${fileName}: unknown tag(s) ${[...unknownTags].sort().join(", ")}`
    );
  }

  return {
    fileName,
    manual,
    readers,
    averages: {
      whyLaw: averageScores(whyLawScores),
      thrive: averageScores(thriveScores),
      contribute: averageScores(contributeScores),
      know: averageScores(knowScores),
    },
    bands: {
      reach: uniqueNonEmpty(rows.map((row) => mapBandValue(row.Reach, fileName, "Reach"))),
      target: uniqueNonEmpty(rows.map((row) => mapBandValue(row.Target, fileName, "Target"))),
      safety: uniqueNonEmpty(rows.map((row) => mapBandValue(row.Safety, fileName, "Safety"))),
    },
    activeTags,
    tagReaderMap,
  };
}

function renderStarSvg(isFilled) {
  return `<svg class="star ${isFilled ? "filled" : "empty"}" viewBox="0 0 24 24" aria-hidden="true">
    <polygon points="12,2 15.3,8.8 22.8,9.8 17.3,15.1 18.7,22.5 12,18.8 5.3,22.5 6.7,15.1 1.2,9.8 8.7,8.8"></polygon>
  </svg>`;
}

function renderStarRow(label, averageValue) {
  if (averageValue === null) {
    return `<div class="rating-row"><div class="label">${escapeHtml(label)}</div><div class="small">No rating submitted</div></div>`;
  }

  const rounded = Math.max(1, Math.min(5, Math.round(averageValue)));
  const stars = [1, 2, 3, 4, 5].map((idx) => renderStarSvg(idx <= rounded)).join("");
  return `<div class="rating-row">
    <div class="label">${escapeHtml(label)}</div>
    <div class="stars-wrap">
      <div class="stars">${stars}</div>
      <span class="rating-number">${averageValue.toFixed(1)} / 5</span>
    </div>
  </div>`;
}

function renderBandList(label, values) {
  if (!values.length) {
    return `<div class="row"><div class="label">${escapeHtml(label)}</div><div class="value">No submissions</div></div>`;
  }

  return `<div class="row">
    <div class="label">${escapeHtml(label)}</div>
    <ul class="band-list">${values
      .map((v) => `<li>${escapeHtml(BAND_LABELS[v] || v)}</li>`)
      .join("")}</ul>
  </div>`;
}

function renderTagBadges(readerLabels) {
  if (!readerLabels.length) return "";
  return `<span class="tag-badges">${readerLabels
    .map((label) => `<span class="tag-badge">${escapeHtml(label.replace("Reader ", ""))}</span>`)
    .join("")}</span>`;
}

function renderTagGrid(activeTags, tagReaderMap) {
  return `<div class="tag-grid">
    ${TAG_DEFINITIONS.map((tag) => {
      const isActive = activeTags.has(tag.name);
      const readerLabels = isActive
        ? [...(tagReaderMap.get(tag.name) || new Set())].sort(
            (a, b) => Number(a.replace("Reader ", "")) - Number(b.replace("Reader ", ""))
          )
        : [];
      const className = isActive
        ? tag.polarity === "positive"
          ? "tag-pill active-positive"
          : "tag-pill active-negative"
        : "tag-pill inactive";
      return `<div class="${className}">${escapeHtml(tag.name)}${renderTagBadges(readerLabels)}</div>`;
    }).join("")}
  </div>`;
}

function renderCompactStars(averageValue) {
  const rounded =
    averageValue === null ? 0 : Math.max(1, Math.min(5, Math.round(averageValue)));
  const stars = [1, 2, 3, 4, 5].map((idx) => renderStarSvg(idx <= rounded)).join("");
  return `<span class="compact-stars">${stars}</span>`;
}

function renderBandRow(label, className, values) {
  const order = ["T14", "T25", "T50", "T100", "T100+"];
  return `<div class="band-row">
    <div class="band-name ${className}">${escapeHtml(label)}</div>
    ${order
      .map((value) => {
        const active = values.includes(value) ? " active" : "";
        return `<div class="band-chip${active}">${escapeHtml(value)}</div>`;
      })
      .join("")}
  </div>`;
}

function renderTagGridFourColumns(activeTags, tagReaderMap) {
  return `<div class="design-tag-grid">
    ${TAG_DEFINITIONS.map((tag) => {
      const isActive = activeTags.has(tag.name);
      const readerLabels = isActive
        ? [...(tagReaderMap.get(tag.name) || new Set())].sort(
            (a, b) => Number(a.replace("Reader ", "")) - Number(b.replace("Reader ", ""))
          )
        : [];
      const className = isActive
        ? tag.polarity === "positive"
          ? "tag-pill active-positive"
          : "tag-pill active-negative"
        : "tag-pill inactive";
      return `<div class="${className}">${escapeHtml(tag.name)}${renderTagBadges(readerLabels)}</div>`;
    }).join("")}
  </div>`;
}

function renderReaders(readers) {
  return readers
    .map(
      (reader) => `<article class="reader-card">
        <h3 class="reader-title">${escapeHtml(reader.label)}</h3>
        <div class="label">Notes</div>
        <div class="notes-box">${escapeHtml(reader.notes)}</div>
        <div class="label">Anything Else?</div>
        <div class="notes-box">${escapeHtml(reader.anythingElse)}</div>
      </article>`
    )
    .join("");
}

function renderStudentDocument(report) {
  return `
    <section class="page summary-page">
      <div class="hero-head">
        <div class="brand-left"><span class="badge">7Sage</span>Admissions</div>
        <div class="brand-right">Committee Review</div>
      </div>
      <div class="summary-banner">Summary</div>
      <div class="section-block readers-section">
        <div class="rail-label">Readers</div>
        <div class="section-body">
          <div class="fit-content fit-readers">
          <div class="reader-grid">
            ${report.readers
              .map(
                (reader) => `<div class="reader-col">
                  <div class="avatar">${escapeHtml(reader.label.replace("Reader ", "R"))}</div>
                  <h3 class="reader-name">${escapeHtml(reader.rawReviewer || reader.label)}</h3>
                  <p class="reader-bio">Lorem ipsum dolor sit amet, consectetur adipiscing elit. Sed do eiusmod tempor incididunt.</p>
                </div>`
              )
              .join("")}
          </div>
          </div>
        </div>
      </div>
      <div class="section-block key-info-section">
        <div class="rail-label">Key Info</div>
        <div class="section-body">
          <div class="fit-content fit-key-info">
          <div class="key-card">
            <div class="key-top">
              <div class="key-top-item"><span class="key-top-label">LSAT:</span> <span class="key-top-value">${escapeHtml(report.manual.lsat || "—")}</span></div>
              <div class="key-top-item"><span class="key-top-label">GPA:</span> <span class="key-top-value">${escapeHtml(report.manual.gpa || "—")}</span></div>
              <div class="key-top-item"><span class="key-top-label">Other:</span> <span class="key-top-value">${escapeHtml(report.manual.otherText || "—")}</span></div>
            </div>
            <div class="metric-grid">
              <div class="metric-col">
                <h4 class="metric-title">Why Law</h4>
                ${renderCompactStars(report.averages.whyLaw)}
              </div>
              <div class="metric-col">
                <h4 class="metric-title">Potential</h4>
                ${renderCompactStars(report.averages.thrive)}
              </div>
              <div class="metric-col">
                <h4 class="metric-title">Contribution</h4>
                ${renderCompactStars(report.averages.contribute)}
              </div>
              <div class="metric-col">
                <h4 class="metric-title">Personality</h4>
                ${renderCompactStars(report.averages.know)}
              </div>
            </div>
            ${renderBandRow("Reach", "reach", report.bands.reach)}
            ${renderBandRow("Target", "target", report.bands.target)}
            ${renderBandRow("Safety", "safety", report.bands.safety)}
          </div>
          </div>
        </div>
      </div>
      <div class="section-block tags-section">
        <div class="rail-label">Tags</div>
        <div class="section-body">
          <div class="fit-content fit-tags">
          ${renderTagGridFourColumns(report.activeTags, report.tagReaderMap)}
          </div>
        </div>
      </div>
    </section>

    <section class="page">
      <header class="page-header">
        <h1 class="doc-title">Committee Review Summary</h1>
        <p class="subtitle">Page 2 of 4 - School Band Summary</p>
      </header>
      ${renderBandList("Reach", report.bands.reach)}
      ${renderBandList("Target", report.bands.target)}
      ${renderBandList("Safety", report.bands.safety)}
      <p class="small">Band values shown are the distinct entries across all three readers.</p>
    </section>

    <section class="page">
      <header class="page-header">
        <h1 class="doc-title">Committee Review Summary</h1>
        <p class="subtitle">Page 3 of 4 - Tag Activation (40 Tags)</p>
      </header>
      ${renderTagGrid(report.activeTags, report.tagReaderMap)}
      <p class="small">Active if selected by at least one reader. Green = positive, red = negative.</p>
    </section>

    <section class="page">
      <header class="page-header">
        <h1 class="doc-title">Committee Review Summary</h1>
        <p class="subtitle">Page 4 of 4 - Reader Comments</p>
      </header>
      ${renderReaders(report.readers)}
    </section>
  `;
}

function getDefaultStudentInput() {
  return {
    lsat: "",
    gpa: "",
    kjd: "Not KJD",
    urm: "Non-URM",
  };
}

function setManualControlsEnabled(enabled) {
  lsatInput.disabled = !enabled;
  gpaInput.disabled = !enabled;
  kjdSelect.disabled = !enabled;
  urmSelect.disabled = !enabled;
}

function getSelectedFile() {
  return String(studentSelect.value || "").trim();
}

function ensureStudentInput(fileName) {
  if (!state.studentInputByFile[fileName]) {
    state.studentInputByFile[fileName] = getDefaultStudentInput();
  }
  return state.studentInputByFile[fileName];
}

function syncManualFormFromSelection() {
  const fileName = getSelectedFile();
  const hasSelection = Boolean(fileName);
  setManualControlsEnabled(hasSelection);
  if (!hasSelection) {
    lsatInput.value = "";
    gpaInput.value = "";
    kjdSelect.value = "Not KJD";
    urmSelect.value = "Non-URM";
    inputHintEl.textContent = "";
    return;
  }

  const manual = ensureStudentInput(fileName);
  lsatInput.value = manual.lsat;
  gpaInput.value = manual.gpa;
  kjdSelect.value = manual.kjd;
  urmSelect.value = manual.urm;
  validateManualInputs();
}

function setStudentSelectorOptions() {
  studentSelect.innerHTML = "";
  if (!state.availableStudents.length) {
    studentSelect.innerHTML = "<option>No valid student names found</option>";
    studentSelect.disabled = true;
    setManualControlsEnabled(false);
    return;
  }

  state.availableStudents.forEach((fileName) => {
    const option = document.createElement("option");
    option.value = fileName;
    option.textContent = fileName;
    studentSelect.appendChild(option);
  });
  studentSelect.disabled = false;
  syncManualFormFromSelection();
}

function renderPreviewFromCurrent() {
  const report = state.currentReport;
  if (!report) {
    previewRoot.innerHTML =
      '<p class="placeholder">No document available. Generate documents first.</p>';
    return;
  }
  previewRoot.innerHTML = `<div class="doc-shell">${renderStudentDocument(report)}</div>`;
  fitAllSummarySections(previewRoot);
  requestAnimationFrame(() => fitAllSummarySections(previewRoot));
}

function validateManualInputs() {
  const warnings = [];
  const lsat = lsatInput.value.trim();
  const gpa = gpaInput.value.trim();

  if (lsat && !/^\d+$/.test(lsat)) {
    warnings.push("LSAT should usually be digits only.");
  }
  if (gpa && !/^\d+(\.\d+)?$/.test(gpa)) {
    warnings.push("GPA should usually look like a decimal number.");
  }

  inputHintEl.textContent = warnings.join(" ");
}

function onManualInputChange() {
  const fileName = getSelectedFile();
  if (!fileName) return;
  const manual = ensureStudentInput(fileName);
  manual.lsat = lsatInput.value.trim();
  manual.gpa = gpaInput.value.trim();
  manual.kjd = kjdSelect.value;
  manual.urm = urmSelect.value;
  validateManualInputs();
}

function fitSectionContent(section) {
  const container = section.querySelector(".section-body");
  const content = container?.querySelector(".fit-content");
  if (!container || !content) return;

  content.style.transform = "scale(1)";
  content.style.width = "100%";

  const availableHeight = container.clientHeight;
  const availableWidth = container.clientWidth;
  const neededHeight = content.scrollHeight;
  const neededWidth = content.scrollWidth;

  if (!availableHeight || !availableWidth || !neededHeight || !neededWidth) return;

  const heightScale = availableHeight / neededHeight;
  const widthScale = availableWidth / neededWidth;
  const scale = Math.min(1, heightScale, widthScale);

  if (scale < 1) {
    content.style.transform = `scale(${scale})`;
    content.style.width = `${100 / scale}%`;
  }
}

function fitAllSummarySections(root = document) {
  root.querySelectorAll(".summary-page .section-block").forEach((section) => {
    fitSectionContent(section);
  });
}

function onWindowResize() {
  if (fitResizeScheduled) return;
  fitResizeScheduled = true;
  requestAnimationFrame(() => {
    fitResizeScheduled = false;
    fitAllSummarySections(previewRoot);
  });
}

async function onCsvSelected(event) {
  clearGeneratedResults({ resetSelector: true });
  generateBtn.disabled = true;
  state.parsedRows = [];
  state.groupedRows = new Map();
  state.availableStudents = [];
  state.studentInputByFile = {};
  state.fileName = "";
  setManualControlsEnabled(false);
  inputHintEl.textContent = "";

  const file = event.target.files[0];
  if (!file) {
    setStatus("Waiting for CSV upload.");
    showValidationMessages();
    return;
  }

  try {
    const csvText = await file.text();
    const { headers, rows } = readCsvRows(csvText);
    const missingHeaders = validateHeaders(headers);
    const { groups, blankFileRows } = groupRowsByFile(rows);

    state.fileName = file.name;
    state.parsedRows = rows;
    state.groupedRows = groups;
    state.availableStudents = [...groups.keys()].sort((a, b) => a.localeCompare(b));
    state.errors = missingHeaders.length
      ? [`Missing required header(s): ${missingHeaders.join(", ")}`]
      : [];
    state.warnings = [];
    if (blankFileRows > 0) {
      state.warnings.push(`Skipped ${blankFileRows} row(s) with blank File.`);
    }

    generateBtn.disabled = state.errors.length > 0 || state.availableStudents.length === 0;
    setStudentSelectorOptions();
    setStatus(
      state.errors.length
        ? `Loaded ${rows.length} row(s), but CSV contract validation failed.`
        : `Loaded ${rows.length} row(s) from ${file.name}. Found ${state.availableStudents.length} student(s).`
    );
  } catch (error) {
    state.errors = [`Could not parse CSV: ${error.message}`];
    generateBtn.disabled = true;
    setStatus("CSV load failed.");
  }

  showValidationMessages();
}

function onGenerateDocuments() {
  clearGeneratedResults();
  if (!state.parsedRows.length) {
    state.errors = ["No CSV rows available. Upload a file first."];
    setStatus("Generation blocked.");
    showValidationMessages();
    return;
  }

  const selectedFile = getSelectedFile();
  if (!selectedFile) {
    state.errors = ["Select a student before generating."];
    setStatus("Generation blocked.");
    showValidationMessages();
    return;
  }

  const rows = state.groupedRows.get(selectedFile);
  if (!rows) {
    state.errors = [`Selected student "${selectedFile}" was not found in CSV data.`];
    setStatus("Generation blocked.");
    showValidationMessages();
    return;
  }

  if (rows.length !== 3) {
    state.errors = [`${selectedFile} has ${rows.length} review(s); requires exactly 3.`];
    setStatus("Generation blocked.");
    showValidationMessages();
    return;
  }

  onManualInputChange();
  const manual = ensureStudentInput(selectedFile);
  const mergedManual = {
    lsat: manual.lsat,
    gpa: manual.gpa,
    kjd: manual.kjd,
    urm: manual.urm,
    otherText: `${manual.kjd} • ${manual.urm}`,
  };

  const report = buildStudentReport(selectedFile, rows, mergedManual);
  state.currentReport = report;
  state.reports = [report];
  renderPreviewFromCurrent();
  downloadPdfBtn.disabled = false;
  printCurrentBtn.disabled = false;
  setStatus(`Generated report for ${selectedFile}.`);
  showValidationMessages();
}

function onStudentSelectionChange() {
  syncManualFormFromSelection();
  state.currentReport = null;
  state.reports = [];
  downloadPdfBtn.disabled = true;
  printCurrentBtn.disabled = true;
  previewRoot.innerHTML =
    '<p class="placeholder">Update metadata as needed, then click Generate Report.</p>';
}

function openPrintWindow(title, bodyHtml) {
  const printWindow = window.open("", "_blank");
  if (!printWindow) {
    state.errors = ["Pop-up blocked. Allow pop-ups to print."];
    showValidationMessages();
    return;
  }

  const printHtml = `<!doctype html>
<html lang="en">
  <head>
    <meta charset="utf-8" />
    <title>${escapeHtml(title)}</title>
    <style>${PRINT_CSS}</style>
  </head>
  <body>
    ${bodyHtml}
    <script>${PRINT_FIT_SCRIPT}</script>
  </body>
</html>`;

  printWindow.document.open();
  printWindow.document.write(printHtml);
  printWindow.document.close();
  printWindow.focus();

  try {
    setTimeout(() => {
      try {
        const bodyContent = printWindow.document?.body?.innerHTML?.trim() || "";
        if (!bodyContent) {
          state.warnings = [
            ...state.warnings,
            "Print window failed to render. Check popup settings and retry.",
          ];
          showValidationMessages();
          return;
        }
        printWindow.print();
      } catch (_error) {
        state.warnings = [
          ...state.warnings,
          "Print fallback encountered an issue. Retry printing if needed.",
        ];
        showValidationMessages();
      }
    }, 250);
  } catch (_error) {
    state.warnings = [
      ...state.warnings,
      "Could not trigger print fallback automatically. You can print manually from the popup.",
    ];
    showValidationMessages();
  }
}

function sanitizeFileName(value) {
  return String(value || "student-report")
    .replace(/[<>:"/\\|?*\x00-\x1F]/g, "")
    .replace(/\s+/g, " ")
    .trim()
    .slice(0, 120);
}

async function onDownloadCurrentStudentPdf() {
  const report = state.currentReport;
  if (!report) return;

  if (typeof window.html2pdf !== "function") {
    state.warnings = [
      ...state.warnings,
      "PDF library is unavailable. Use Print Current Student as fallback.",
    ];
    showValidationMessages();
    return;
  }

  const sourceDoc = previewRoot.querySelector(".doc-shell");
  if (!sourceDoc) {
    state.warnings = [
      ...state.warnings,
      "Could not find preview content. Generate the report again and retry PDF.",
    ];
    showValidationMessages();
    return;
  }

  fitAllSummarySections(previewRoot);
  await new Promise((resolve) => requestAnimationFrame(resolve));

  const staging = document.createElement("div");
  staging.style.position = "fixed";
  staging.style.left = "0";
  staging.style.top = "0";
  staging.style.width = "612px";
  staging.style.opacity = "0";
  staging.style.pointerEvents = "none";
  staging.style.zIndex = "-1";
  staging.style.overflow = "hidden";

  const clone = sourceDoc.cloneNode(true);
  clone.classList.add("pdf-export-root");
  clone.style.width = "612px";
  clone.style.margin = "0";
  clone.style.padding = "0";
  staging.appendChild(clone);
  document.body.appendChild(staging);

  try {
    if (document.fonts?.ready) {
      await document.fonts.ready;
    }
    fitAllSummarySections(staging);
    await new Promise((resolve) => requestAnimationFrame(resolve));
    fitAllSummarySections(staging);
    await new Promise((resolve) => requestAnimationFrame(resolve));

    const filename = `${sanitizeFileName(report.fileName) || "student-report"}.pdf`;
    const options = {
      margin: 0,
      filename,
      image: { type: "jpeg", quality: 0.98 },
      html2canvas: {
        scale: 2,
        useCORS: true,
        backgroundColor: "#ffffff",
        width: 612,
        windowWidth: 612,
        x: 0,
        y: 0,
        scrollX: 0,
        scrollY: 0,
      },
      jsPDF: { unit: "px", format: [612, 792], orientation: "portrait" },
      pagebreak: { mode: ["css"] },
    };
    await window.html2pdf().set(options).from(clone).save();
  } catch (_error) {
    state.warnings = [
      ...state.warnings,
      "PDF export failed. Use Print Current Student as fallback.",
    ];
    showValidationMessages();
  } finally {
    staging.remove();
  }
}

function onPrintCurrentStudent() {
  const report = state.currentReport;
  if (!report) return;
  openPrintWindow(report.fileName, `<div class="doc-shell">${renderStudentDocument(report)}</div>`);
}
