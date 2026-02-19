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
  "Softs",
  "Notes",
  "Anything Else?",
];

const RUBRIC_DELTA = {
  "Strongly Agree": 1,
  Agree: 0,
  Neutral: -1,
  Disagree: -2,
  "Strongly Disagree": -2,
};

const BAND_ORDER = ["T3", "T6", "T14", "T20", "T30", "T50", "T75", "T100", "T100+"];

const BAND_LABELS = {
  T3: "Top 3",
  T6: "Top 6",
  T14: "Top 14",
  T20: "Top 20",
  T30: "Top 30",
  T50: "Top 50",
  T75: "Top 75",
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
const TAG_POLARITY_MAP = new Map(TAG_DEFINITIONS.map((tag) => [tag.name, tag.polarity]));
const READER_PROFILES = [
  {
    fullName: "Brigitte Suhr",
    firstName: "Brigitte",
    aliases: ["Brigitte", "Briggite"],
    headshotUrl:
      "https://www.gravatar.com/avatar/759b37b8091b593596f86f07072fe396?size=320&default=robohash",
    bio: "",
  },
  {
    fullName: "Sam Riley",
    firstName: "Sam",
    aliases: ["Sam Riley"],
    headshotUrl: "",
    bio: "",
  },
  {
    fullName: "Sam Kwak",
    firstName: "Sam",
    aliases: ["Sam Kwak"],
    headshotUrl: "",
    bio: "",
  },
  {
    fullName: "Reyes Aguilar",
    firstName: "Reyes",
    aliases: ["Reyes Aguilar", "Reyes"],
    headshotUrl:
      "https://www.gravatar.com/avatar/5df779a3d2aef8fbdd29088b8cc485d2?size=320&default=robohash",
    bio: "",
  },
  {
    fullName: "Jennifer Kott",
    firstName: "Jennifer",
    aliases: ["Jennifer Kott", "Jennifer"],
    headshotUrl:
      "https://www.gravatar.com/avatar/29d0502c4ea11721ff29d3d1fa1c3bdd?size=320&default=robohash",
    bio:
      "Lorem ipsum dolor sit amet, consectetur adipiscing elit. Ut tincidunt erat et quam lobortis porttitor. Proin eleifend nisi in neque commodo efficitur. Quisque fermentum mi.",
  },
];

const TAG_FONT_BASE_PX = 9;
const TAG_FONT_MIN_PX = 7;
const NOTES_FONT_BASE_PX = 10;
const NOTES_FONT_MIN_PX = 5;
const NOTES_LINE_BASE = 1.3;
const NOTES_LINE_TIGHT = 1.15;
const NOTES_LINE_TIGHTER = 1.05;
const NOTES_LINE_MIN = 1.0;

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
  *, *::before, *::after { box-sizing: border-box; }
  @page { size: 8.5in 11in; margin: 0; }
  body { margin: 0; font-family: "Lexend", "Segoe UI", Tahoma, sans-serif; color: #202530; background: #fffffe; }
  .doc-shell { background: #fff; }
  .page { width: 612px; min-height: 792px; height: 792px; padding: 24px; margin: 0; break-after: page; background: #fffffe; }
  .page.summary-page { padding: 0; min-height: 792px; height: 792px; width: 612px; transform: scale(1.3333333333); transform-origin: top left; margin-bottom: 264px; background: #fffffe; }
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
  .star.half polygon { stroke: #b78600; }
  .waiting-page { display: flex; align-items: center; justify-content: center; }
  .waiting-copy { font-family: "Fraunces", "Times New Roman", serif; font-size: 36px; font-weight: 700; color: #98a2b3; text-transform: lowercase; }
  .band-list { margin: 0; padding-left: 18px; }
  .band-list li { margin-bottom: 6px; }
  .tag-grid { display: grid; grid-template-columns: repeat(3, minmax(0, 1fr)); gap: 8px; }
  .tag-pill { --tag-font-size: 9px; border: 1px solid #e0e6f2; border-radius: 999px; padding: 2px 6px; font-size: var(--tag-font-size); line-height: 1.1; font-weight: 600; text-align: center; display: flex; align-items: center; justify-content: center; position: relative; overflow: visible; color: #4b5563; }
  .tag-text { display: -webkit-box; width: 100%; overflow: hidden; -webkit-line-clamp: 2; -webkit-box-orient: vertical; }
  .tag-pill.active-positive { background: #ecfdf3; border-color: #86efac; color: #166534; }
  .tag-pill.active-negative { background: #fef2f2; border-color: #fecaca; color: #7f1d1d; }
  .tag-badges { display: inline-flex; gap: 4px; position: absolute; top: -12px; right: 8px; margin-left: 0; vertical-align: baseline; z-index: 2; }
  .tag-badge { min-width: 14px; height: 14px; border-radius: 999px; display: inline-flex; align-items: center; justify-content: center; font-size: 9px; font-weight: 700; background: #111827; color: #fff; border: 1px solid #fff; }
  .tag-pill.inactive { color: #5e6778; background: #fff; }
  .reader-card { border: 1px solid #d8dee9; border-radius: 8px; padding: 10px; margin-bottom: 10px; }
  .reader-title { margin: 0 0 8px; font-size: 14px; }
  .notes-box { border: 1px solid #333; min-height: 1.5in; padding: 10px; white-space: pre-wrap; margin-bottom: 8px; }
  .summary-banner { height: 58px; background: linear-gradient(-45deg, #15b79e 0%, #227f9c 100%); color: #fcfaf8; text-align: center; font-family: "Fraunces", "Times New Roman", serif; font-size: 30px; font-weight: 700; line-height: 58px; margin-bottom: 0; }
  .section-block { display: grid; grid-template-columns: 1fr; gap: 0; margin-bottom: 0; width: 100%; background: #fff; border-top: 0; border-bottom: 0; }
  .summary-page .section-block { background: #fffffe; }
  .section-block.basics-section { height: 52px; }
  .section-block.readers-section { height: 150px; margin-top: -8px; }
  .section-block.readers-section .section-body { height: 100%; overflow: hidden; padding: 4px 8px; }
  .section-block.readers-section .avatar { width: 36px; height: 36px; border-radius: 9px; font-size: 11px; }
  .section-block.readers-section .reader-name { font-size: 13px; }
  .section-block.readers-section .reader-bio { font-size: 10px; line-height: 1.2; -webkit-line-clamp: 4; }
  .fit-readers { display: block; }
  .readers-card { border: 2px solid #d8dee9; border-radius: 12px; padding: 4px 6px; height: calc(100% - 14px); }
  .section-block.basics-section .section-body { height: 100%; padding: 4px 10px; }
  .section-block.key-takeaways-section { height: 232px; margin-top: 16px; }
  .section-block.key-takeaways-section .section-body { height: 100%; padding: 6px 12px; }
  .section-block.tags-section { height: 300px; border-bottom: 0; }
  .section-block.tags-section .section-body { height: 100%; overflow: hidden; }
  .section-block.tags-section .section-body { padding: 6px 8px; }
  .fit-tags { display: flex; flex-direction: column; height: 100%; }
  .tags-card { border: 2px solid #d8dee9; border-radius: 12px; padding: 6px 8px; flex: 1; display: flex; flex-direction: column; }
  .tags-grid-wrap { flex: 1; display: flex; align-items: center; justify-content: center; min-height: 0; width: 100%; }
  .fit-tags .design-tag-grid { transform: scale(0.93); transform-origin: center; width: 100%; height: 100%; }
  .basics-grid { display: flex; align-items: center; justify-content: center; gap: 10px; flex-wrap: nowrap; }
  .basics-item { display: inline-flex; gap: 4px; align-items: baseline; text-align: left; }
  .basics-label { font-size: 10px; font-weight: 700; letter-spacing: 0.12em; text-transform: uppercase; color: #475467; }
  .basics-value { font-size: 12px; font-weight: 700; }
  .basics-divider { font-size: 12px; color: #98a2b3; padding: 0 4px; }
  .takeaways-row { display: grid; gap: 8px; margin-bottom: 8px; }
  .takeaways-row-top { grid-template-columns: repeat(3, minmax(0, 1fr)); }
  .takeaways-row-bottom { grid-template-columns: repeat(2, minmax(0, 1fr)); }
  .takeaway-item { text-align: center; display: grid; gap: 4px; }
  .takeaway-title { font-family: "Fraunces", "Times New Roman", serif; font-size: 10px; font-weight: 700; text-transform: uppercase; letter-spacing: 0.1em; color: #475467; }
  .takeaway-softs-value { font-size: 14px; font-weight: 700; }
  .takeaways-bands .band-row { font-size: 9px; gap: 0; padding: 3px 6px; grid-template-columns: 52px repeat(9, minmax(0, 1fr)); }
  .takeaways-card { border: 2px solid #d8dee9; border-radius: 12px; padding: 8px 10px; display: grid; gap: 8px; }
  .section-title { font-family: "Fraunces", "Times New Roman", serif; text-transform: uppercase; letter-spacing: 0.22em; font-size: 10px; font-weight: 700; color: #94a3b8; text-align: center; margin: 0 0 4px; }
  .section-block.readers-section .section-title { margin-bottom: 2px; }
  .rail-label { writing-mode: vertical-rl; transform: rotate(180deg); background: #334155; color: #fff; width: 40px; border-radius: 18px 0 0 18px; text-align: center; font-family: "Fraunces", "Times New Roman", serif; font-size: 22px; font-weight: 700; letter-spacing: 0.4px; padding: 16px 8px; }
  .section-body { border: 0; border-radius: 0; padding: 14px 16px; background: transparent; width: 100%; min-width: 0; overflow: hidden; }
  .fit-content { transform-origin: top left; width: 100%; }
  .reader-grid { display: grid; grid-template-columns: repeat(3, minmax(0, 1fr)); grid-template-rows: 1fr; gap: 10px; height: 100%; align-content: stretch; }
  .reader-col { padding: 2px 0; position: relative; display: flex; flex-direction: column; align-items: center; gap: 4px; min-height: 0; }
  .summary-page .reader-col:not(:last-child)::after { content: ""; position: absolute; top: 10%; bottom: 10%; right: -6px; width: 2px; background: #d8dee9; }
  .avatar { width: 52px; height: 52px; border-radius: 12px; margin: 0; display: flex; align-items: center; justify-content: center; background: linear-gradient(135deg, #2b7abf, #14b8a6); color: #fff; font-weight: 800; font-size: 14px; }
  .avatar.has-photo { background: transparent; padding: 0; }
  .avatar img { width: 100%; height: 100%; border-radius: inherit; object-fit: cover; display: block; }
  .reader-name { text-align: center; margin: 0 0 1px; font-size: 16px; font-family: "Fraunces", "Times New Roman", serif; white-space: nowrap; text-overflow: ellipsis; overflow: hidden; }
  .reader-bio { margin: 0; text-align: center; font-size: 11px; line-height: 1.3; display: -webkit-box; -webkit-line-clamp: 4; -webkit-box-orient: vertical; overflow: hidden; }
  .key-card { border: 0; border-radius: 0; overflow: hidden; width: 100%; height: 100%; display: flex; flex-direction: column; }
  .key-top { display: grid; grid-template-columns: minmax(0, 0.8fr) minmax(0, 0.8fr) minmax(0, 1.4fr) minmax(0, 0.9fr); border-bottom: 2px solid #d8dee9; }
  .key-top-item { padding: 8px 12px; font-size: 19px; font-weight: 700; white-space: nowrap; position: relative; }
  .key-top-label { opacity: 0.9; }
  .key-top-value { margin-left: 4px; font-weight: 800; }
  .key-top-item.other-item .key-top-value { font-size: 13px; font-weight: 500; letter-spacing: 0.1px; }
  .key-top-item.softs-item .key-top-value { font-weight: 500; }
  .key-top-item:not(:last-child)::after { content: ""; position: absolute; top: 50%; right: 0; transform: translateY(-50%); width: 2px; height: 60%; background: #d8dee9; }
  .metric-grid { display: grid; grid-template-columns: repeat(4, minmax(0, 1fr)); border-bottom: 2px solid #d8dee9; }
  .metric-col { padding: 8px 10px; text-align: center; position: relative; }
  .metric-col:not(:last-child)::after { content: ""; position: absolute; top: 50%; right: 0; transform: translateY(-50%); width: 2px; height: 60%; background: #d8dee9; }
  .metric-title { margin: 0 0 6px; font-size: 20px; font-weight: 700; }
  .compact-stars { display: inline-flex; gap: 3px; justify-content: center; }
  .compact-stars .star { width: 14px; height: 14px; }
  .band-row { display: grid; grid-template-columns: 72px repeat(9, minmax(0, 1fr)); align-items: center; gap: 0; padding: 5px 8px; font-size: 11px; position: relative; flex: 1 0 0; }
  .band-name { font-weight: 700; border-radius: 999px; text-align: center; padding: 2px 6px; margin-right: 6px; position: relative; z-index: 3; }
  .band-name.reach { background: #fff2df; color: #b45309; }
  .band-name.target { background: #e0f2fe; color: #0c4a6e; }
  .band-name.safety { background: #dcfce7; color: #14532d; }
  .band-row.row-reach .band-chip.in-range { color: #b45309; }
  .band-row.row-target .band-chip.in-range { color: #0c4a6e; }
  .band-row.row-safety .band-chip.in-range { color: #14532d; }
  .band-range { grid-row: 1; border-radius: 8px; min-height: 18px; align-self: stretch; z-index: 1; }
  .band-range.range-reach { background: #fff2df; border: 1px solid #f3ce73; }
  .band-range.range-target { background: #e0f2fe; border: 1px solid #7dd3fc; }
  .band-range.range-safety { background: #dcfce7; border: 1px solid #86efac; }
  .band-chip { display: flex; align-items: center; justify-content: center; text-align: center; padding: 2px 0; min-height: 18px; font-weight: 700; color: #111827; position: relative; z-index: 2; }
  .design-tag-grid { display: grid; grid-template-columns: repeat(5, minmax(0, 1fr)); grid-template-rows: repeat(8, minmax(0, 1fr)); gap: 6px 8px; height: 100%; align-content: stretch; flex: 1; min-height: 0; }
  .page.reader-page { padding: 16px; display: grid; grid-template-rows: 81% 19%; gap: 8px; background: #fffffe; }
  .reader-page-main { min-height: 0; display: grid; grid-template-rows: repeat(3, 1fr); gap: 0; }
  .reader-page-main .reader-section { background: #fffffe; }
  .reader-page-footer { display: grid; grid-template-columns: repeat(2, minmax(0, 1fr)); gap: 8px; min-height: 0; position: relative; background: #fffffe; margin: 0 -16px; padding: 0 16px; }
  .reader-page-footer::before { content: ""; position: absolute; left: 0; right: 0; top: -6px; height: 2px; background: #d8dee9; }
  .reader-footer-box { border: 1px solid #d8dee9; border-radius: 8px; padding: 6px 8px; display: grid; grid-template-rows: auto 1fr; gap: 4px; overflow: hidden; font-size: 10px; line-height: 1.3; }
  .reader-footer-title { font-family: "Fraunces", "Times New Roman", serif; font-weight: 700; font-size: 9px; text-transform: uppercase; letter-spacing: 0.12em; color: #475467; }
  .reader-footer-body { white-space: pre-wrap; }
  .reader-sections { display: grid; grid-template-rows: repeat(3, 1fr); gap: 12px; height: 100%; }
  .reader-section { border: 2px solid #d8dee9; border-radius: 12px; padding: 8px 10px; margin: 0; display: flex; flex-direction: column; min-height: 0; position: relative; }
  .reader-section-title { font-family: "Fraunces", "Times New Roman", serif; font-weight: 700; font-size: 9px; letter-spacing: 0.18em; text-transform: uppercase; color: #94a3b8; margin-bottom: 4px; text-align: center; }
  .reader-rows { display: grid; grid-template-rows: 40% 60%; height: 100%; gap: 6px; min-height: 0; }
  .reader-row-top { display: grid; grid-template-columns: 0.95fr 0.75fr 1.3fr; gap: 8px; min-height: 0; padding: 0 4px; }
  .reader-page .reader-col { padding: 0; position: relative; display: grid; grid-template-columns: 1fr; align-items: start; }
  .reader-col.ratings, .reader-col.bands { display: grid; grid-template-columns: 1fr; grid-auto-rows: minmax(0, 1fr); gap: 3px; min-height: 0; }
  .reader-col.tags { display: grid; grid-template-columns: repeat(2, minmax(0, 1fr)); gap: 3px; align-content: start; min-height: 0; }
  .reader-rating-row { display: flex; align-items: center; justify-content: space-between; gap: 6px; }
  .reader-rating-label { font-size: 8px; font-weight: 600; text-transform: none; letter-spacing: 0; color: #475467; min-width: 0; flex: 1; }
  .reader-rating-pill { border: 1px solid #e2e8f0; border-radius: 999px; padding: 2px 6px; font-size: 7px; font-weight: 600; line-height: 1.1; background: #f8fafc; color: #334155; white-space: nowrap; }
  .reader-rating-pill.empty { color: #94a3b8; background: #fff; }
  .reader-rating-empty { font-size: 8px; color: #98a2b3; }
  .reader-page .compact-stars .star { width: 12px; height: 12px; }
  .reader-band-row { display: grid; grid-template-columns: auto 1fr; gap: 4px; align-items: center; font-size: 9px; }
  .reader-band-label { font-weight: 700; text-transform: uppercase; letter-spacing: 0.06em; font-size: 8px; color: #475467; }
  .reader-band-value { font-weight: 600; }
  .reader-notes { border: 0; border-radius: 0; padding: 0 4px; --notes-font-size: 10px; --notes-line-height: 1.3; font-size: var(--notes-font-size); line-height: var(--notes-line-height); height: 100%; overflow: hidden; }
  .reader-notes .label { font-weight: 700; font-size: 9px; text-transform: uppercase; letter-spacing: 0.12em; margin-bottom: 2px; color: #475467; }
  .reader-notes-body { margin: 0; white-space: pre-wrap; }
  .reader-tag-pill { border: 1px solid #e0e6f2; border-radius: 999px; padding: 1px 3px; font-size: 7px; line-height: 1.1; text-align: center; display: flex; align-items: center; justify-content: center; min-height: 16px; }
  .reader-tag-pill.active-positive { background: #ecfdf3; border-color: #86efac; color: #166534; }
  .reader-tag-pill.active-negative { background: #fef2f2; border-color: #fecaca; color: #7f1d1d; }
  .reader-tag-pill.inactive { color: #5e6778; background: #fff; }
  .reader-tag-empty { font-size: 10px; color: #98a2b3; align-self: center; }
  .tag-explanation-page { padding: 24px; background: #fffffe; display: flex; flex-direction: column; height: 792px; }
  .tag-explanation-grid { display: grid; grid-template-columns: 1fr; grid-template-rows: repeat(6, minmax(0, 1fr)); gap: 12px; height: 100%; align-content: stretch; flex: 1; }
  .tag-explanation-item { border: 1px solid #d8dee9; border-radius: 8px; padding: 10px 12px; background: transparent; display: grid; gap: 6px; min-height: 0; }
  .tag-explanation-title { display: flex; align-items: center; gap: 8px; }
  .tag-explanation-pill { padding: 4px 10px; font-size: 10px; line-height: 1.1; }
  .tag-explanation-body { font-size: 11px; line-height: 1.4; color: #1f2937; }
  .tag-explanation-empty { display: flex; align-items: center; justify-content: center; height: 100%; font-size: 14px; color: #98a2b3; }
`;

const PRINT_FIT_SCRIPT = `
  (() => {
    const TAG_FONT_BASE_PX = 9;
    const TAG_FONT_MIN_PX = 7;
    const NOTES_FONT_BASE_PX = 10;
    const NOTES_FONT_MIN_PX = 5;
    const NOTES_LINE_BASE = 1.3;
    const NOTES_LINE_TIGHT = 1.15;
    const NOTES_LINE_TIGHTER = 1.05;
    const NOTES_LINE_MIN = 1.0;

    function fitTagFonts(root = document) {
      const pills = root.querySelectorAll(".tag-pill");
      pills.forEach((pill) => {
        const textEl = pill.querySelector(".tag-text");
        if (!textEl) return;

        const prevDisplay = textEl.style.display;
        const prevClamp = textEl.style.webkitLineClamp;
        const prevOrient = textEl.style.webkitBoxOrient;

        textEl.style.display = "block";
        textEl.style.webkitLineClamp = "unset";
        textEl.style.webkitBoxOrient = "initial";

        const setFont = (sizePx) => {
          pill.style.setProperty("--tag-font-size", sizePx + "px");
        };

        const getLineHeight = () => {
          const computed = window.getComputedStyle(textEl);
          const fontSize = parseFloat(computed.fontSize) || TAG_FONT_BASE_PX;
          let lineHeight = parseFloat(computed.lineHeight);
          if (!lineHeight || Number.isNaN(lineHeight)) {
            lineHeight = fontSize * 1.1;
          }
          return { lineHeight, fontSize };
        };

        setFont(TAG_FONT_BASE_PX);
        let { lineHeight } = getLineHeight();
        let maxHeight = lineHeight * 2 + 0.5;
        let fullHeight = textEl.scrollHeight;

        if (fullHeight > maxHeight) {
          let low = TAG_FONT_MIN_PX;
          let high = TAG_FONT_BASE_PX;
          let best = TAG_FONT_MIN_PX;

          while (low <= high) {
            const mid = Math.floor((low + high) / 2);
            setFont(mid);
            ({ lineHeight } = getLineHeight());
            maxHeight = lineHeight * 2 + 0.5;
            fullHeight = textEl.scrollHeight;

            if (fullHeight <= maxHeight) {
              best = mid;
              low = mid + 1;
            } else {
              high = mid - 1;
            }
          }
          setFont(best);
        }

        textEl.style.display = prevDisplay;
        textEl.style.webkitLineClamp = prevClamp;
        textEl.style.webkitBoxOrient = prevOrient;
      });
    }

    function fitReaderNotes(root = document) {
      const notesBlocks = root.querySelectorAll(".reader-notes");
      notesBlocks.forEach((notes) => {
        const body = notes.querySelector(".reader-notes-body");
        if (!body) return;

        const label = notes.querySelector(".label");
        const computed = window.getComputedStyle(notes);
        const paddingTop = parseFloat(computed.paddingTop) || 0;
        const paddingBottom = parseFloat(computed.paddingBottom) || 0;
        const labelHeight = label ? label.offsetHeight : 0;
        const labelMarginBottom = label ? parseFloat(window.getComputedStyle(label).marginBottom) || 0 : 0;

        const availableHeight =
          notes.clientHeight - paddingTop - paddingBottom - labelHeight - labelMarginBottom;
        if (!availableHeight) return;

        const setFont = (sizePx) => {
          notes.style.setProperty("--notes-font-size", sizePx + "px");
        };
        const setLineHeight = (value) => {
          notes.style.setProperty("--notes-line-height", value);
        };

        setLineHeight(NOTES_LINE_BASE);
        setFont(NOTES_FONT_BASE_PX);
        if (body.scrollHeight <= availableHeight) return;

        let low = NOTES_FONT_MIN_PX;
        let high = NOTES_FONT_BASE_PX;
        let best = NOTES_FONT_MIN_PX;

        while (low <= high) {
          const mid = Math.floor((low + high) / 2);
          setFont(mid);
          if (body.scrollHeight <= availableHeight) {
            best = mid;
            low = mid + 1;
          } else {
            high = mid - 1;
          }
        }
        setFont(best);
        setLineHeight(NOTES_LINE_BASE);

        if (body.scrollHeight <= availableHeight) return;
        setLineHeight(NOTES_LINE_TIGHT);
        if (body.scrollHeight <= availableHeight) return;
        setLineHeight(NOTES_LINE_TIGHTER);
        if (body.scrollHeight <= availableHeight) return;
        setLineHeight(NOTES_LINE_MIN);
      });
    }

    function fitSectionContent(section) {
      const container = section.querySelector(".section-body");
      const content = container?.querySelector(".fit-content");
      if (!container || !content) return;

      content.style.transform = "scale(1)";
      content.style.width = "100%";

      const availableHeight = container.clientHeight;
      const neededHeight = content.scrollHeight;
      if (!availableHeight || !neededHeight) return;

      function applyScale(scale) {
        if (scale < 1) {
          content.style.transform = "scale(" + scale + ")";
          content.style.width = (100 / scale) + "%";
        } else {
          content.style.transform = "scale(1)";
          content.style.width = "100%";
        }
      }

      let scale = 1;
      const heightScale = availableHeight / neededHeight;

      // Fill vertical space first.
      if (heightScale < 1) {
        scale = heightScale;
      }

      applyScale(scale);

    }

    function fitAllSummarySections() {
      fitTagFonts(document);
      document.querySelectorAll(".summary-page .section-block").forEach(fitSectionContent);
    }

    window.addEventListener("load", () => {
      fitAllSummarySections();
      fitReaderNotes();
      const finish = () => {
        fitAllSummarySections();
        fitReaderNotes();
        window.focus();
        window.print();
      };
      if (document.fonts?.ready) {
        document.fonts.ready.then(() => setTimeout(finish, 80));
      } else {
        setTimeout(finish, 80);
      }
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
const summaryInput = document.getElementById("summaryInput");
const nextStepsInput = document.getElementById("nextStepsInput");
const inputHintEl = document.getElementById("inputHint");
const downloadPdfBtn = document.getElementById("downloadPdfBtn");
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
summaryInput.addEventListener("input", onManualInputChange);
nextStepsInput.addEventListener("input", onManualInputChange);
downloadPdfBtn.addEventListener("click", onDownloadCurrentStudentPdf);
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

function normalizeProfileKey(value) {
  return String(value || "")
    .trim()
    .toLowerCase()
    .replace(/\s+/g, " ");
}

const READER_PROFILE_NAME_MAP = new Map(
  READER_PROFILES.flatMap((profile) => {
    const keys = [
      profile.fullName,
      ...(profile.aliases || []),
    ]
      .map(normalizeProfileKey)
      .filter(Boolean);
    return keys.map((key) => [key, profile]);
  })
);

const READER_PROFILE_FIRSTNAME_MAP = READER_PROFILES.reduce((acc, profile) => {
  const key = normalizeProfileKey(profile.firstName);
  if (!key) return acc;
  if (!acc.has(key)) acc.set(key, []);
  acc.get(key).push(profile);
  return acc;
}, new Map());

function getReaderProfile(name) {
  const normalized = normalizeProfileKey(name);
  if (!normalized) return null;
  const direct = READER_PROFILE_NAME_MAP.get(normalized);
  if (direct) return direct;

  if (!normalized.includes(" ")) {
    const candidates = READER_PROFILE_FIRSTNAME_MAP.get(normalized) || [];
    if (candidates.length === 1) return candidates[0];
  }
  return null;
}

function getReaderInitials(label) {
  const trimmed = String(label || "").trim();
  if (!trimmed) return "R";
  return trimmed
    .split(/\s+/)
    .map((part) => part[0])
    .slice(0, 2)
    .join("")
    .toUpperCase();
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

function normalizeRubricValue(value, fileName, fieldLabel) {
  const trimmed = String(value || "").trim();
  if (!trimmed) return null;
  if (!(trimmed in RUBRIC_DELTA)) {
    state.warnings.push(`${fileName}: unknown ${fieldLabel} value "${trimmed}"`);
    return null;
  }
  return trimmed;
}

function computeStarsFromRatings(ratings) {
  const filtered = ratings.filter(Boolean);
  if (!filtered.length) return null;
  let stars = 4;
  filtered.forEach((rating) => {
    stars += RUBRIC_DELTA[rating] ?? 0;
  });
  if (stars > 5) stars = 5;
  if (stars < 1) stars = 1;
  return stars;
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
  const normalized = trimmed.toUpperCase().replace(/\s+/g, "");
  const aliasMap = {
    "3": "T3",
    "6": "T6",
    "14": "T14",
    "20": "T20",
    "30": "T30",
    "50": "T50",
    "75": "T75",
    "100": "T100",
    "100+": "T100+",
    T3: "T3",
    T6: "T6",
    T14: "T14",
    T20: "T20",
    T25: "T20",
    T30: "T30",
    T50: "T50",
    T75: "T75",
    T100: "T100",
    "T100+": "T100+",
  };
  const canonical = aliasMap[normalized];
  if (canonical && BAND_LABELS[canonical]) return canonical;
  state.warnings.push(`${fileName}: unknown ${fieldLabel} value "${trimmed}"`);
  return `Unknown band: ${trimmed}`;
}

function mapSoftsValue(value, fileName) {
  const trimmed = String(value || "").trim().toUpperCase();
  if (!trimmed) return null;
  const match = /^T([1-4])$/.exec(trimmed);
  if (!match) {
    state.warnings.push(`${fileName}: unknown Softs value "${trimmed}"`);
    return null;
  }
  return Number(match[1]);
}

function computeSoftsDisplay(rows, fileName) {
  const values = rows
    .map((row) => mapSoftsValue(row.Softs, fileName))
    .filter((value) => value !== null);
  if (!values.length) return null;

  const counts = values.reduce((acc, value) => {
    acc.set(value, (acc.get(value) || 0) + 1);
    return acc;
  }, new Map());

  if (counts.size === 1) {
    const [onlyValue] = counts.keys();
    return `T${onlyValue}`;
  }

  if (counts.size === 2) {
    const sortedCounts = [...counts.entries()].sort((a, b) => b[1] - a[1]);
    return `T${sortedCounts[0][0]}/T${sortedCounts[1][0]}`;
  }

  const average =
    values.reduce((sum, value) => sum + value, 0) / values.length;
  return `T${average.toFixed(1)}`;
}

function uniqueNonEmpty(values) {
  return [...new Set(values.map((v) => String(v || "").trim()).filter(Boolean))];
}

function buildStudentReport(fileName, rows, manual) {
  const unknownTags = new Set();
  const summaryRatings = {
    whyLaw: [],
    thrive: [],
    contribute: [],
    know: [],
  };
  const randomizedRows = shuffleReaders(rows);
  const labeledReaders = randomizedRows.map((row, index) => {
    const rawTags = parseTags(row.Tags);
    const tags = rawTags.filter((tag) => TAG_NAME_SET.has(tag));
    rawTags.forEach((tag) => {
      if (!TAG_NAME_SET.has(tag)) {
        unknownTags.add(tag);
      }
    });
    return {
      row,
      label: `Reader ${index + 1}`,
      tags,
    };
  });

  const normalizeBandDisplay = (value) => {
    const trimmed = String(value || "").trim();
    if (!trimmed || trimmed.startsWith("Unknown band:")) return "—";
    return trimmed;
  };

  const readers = labeledReaders.map(({ row, label, tags }) => {
    const ratingLabels = {
      whyLaw: normalizeRubricValue(row["Why Law?"], fileName, "Why Law?"),
      thrive: normalizeRubricValue(row["Thrive?"], fileName, "Thrive?"),
      contribute: normalizeRubricValue(row["Contribute?"], fileName, "Contribute?"),
      know: normalizeRubricValue(row["Know?"], fileName, "Know?"),
    };
    if (ratingLabels.whyLaw) summaryRatings.whyLaw.push(ratingLabels.whyLaw);
    if (ratingLabels.thrive) summaryRatings.thrive.push(ratingLabels.thrive);
    if (ratingLabels.contribute) summaryRatings.contribute.push(ratingLabels.contribute);
    if (ratingLabels.know) summaryRatings.know.push(ratingLabels.know);

    return {
      label,
      notes: row.Notes || "",
      anythingElse: row["Anything Else?"] || "",
      rawReviewer: row.Reviewer || "",
      ratingLabels,
      ratingStars: {
        whyLaw: computeStarsFromRatings([ratingLabels.whyLaw]),
        thrive: computeStarsFromRatings([ratingLabels.thrive]),
        contribute: computeStarsFromRatings([ratingLabels.contribute]),
        know: computeStarsFromRatings([ratingLabels.know]),
      },
      bands: {
        reach: normalizeBandDisplay(mapBandValue(row.Reach, fileName, "Reach")),
        target: normalizeBandDisplay(mapBandValue(row.Target, fileName, "Target")),
        safety: normalizeBandDisplay(mapBandValue(row.Safety, fileName, "Safety")),
      },
      softs: (() => {
        const softValue = mapSoftsValue(row.Softs, fileName);
        return softValue ? `T${softValue}` : "—";
      })(),
      tags,
    };
  });

  const activeTags = new Set();
  const tagReaderMap = new Map();
  labeledReaders.forEach(({ label, tags }) => {
    tags.forEach((tag) => {
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
    summaryStars: {
      whyLaw: computeStarsFromRatings(summaryRatings.whyLaw),
      thrive: computeStarsFromRatings(summaryRatings.thrive),
      contribute: computeStarsFromRatings(summaryRatings.contribute),
      know: computeStarsFromRatings(summaryRatings.know),
    },
    bands: {
      reach: uniqueNonEmpty(rows.map((row) => mapBandValue(row.Reach, fileName, "Reach"))),
      target: uniqueNonEmpty(rows.map((row) => mapBandValue(row.Target, fileName, "Target"))),
      safety: uniqueNonEmpty(rows.map((row) => mapBandValue(row.Safety, fileName, "Safety"))),
    },
    softsDisplay: computeSoftsDisplay(rows, fileName),
    activeTags,
    tagReaderMap,
  };
}

function renderStarSvg(type, idSuffix) {
  if (type === "half") {
    return `<svg class="star half" viewBox="0 0 24 24" aria-hidden="true">
      <defs>
        <linearGradient id="half-fill-${idSuffix}" x1="0" x2="1" y1="0" y2="0">
          <stop offset="50%" stop-color="#f4b400"></stop>
          <stop offset="50%" stop-color="#eef2f7"></stop>
        </linearGradient>
      </defs>
      <polygon fill="url(#half-fill-${idSuffix})" points="12,2 15.3,8.8 22.8,9.8 17.3,15.1 18.7,22.5 12,18.8 5.3,22.5 6.7,15.1 1.2,9.8 8.7,8.8"></polygon>
    </svg>`;
  }

  return `<svg class="star ${type}" viewBox="0 0 24 24" aria-hidden="true">
    <polygon points="12,2 15.3,8.8 22.8,9.8 17.3,15.1 18.7,22.5 12,18.8 5.3,22.5 6.7,15.1 1.2,9.8 8.7,8.8"></polygon>
  </svg>`;
}

function renderStarRow(label, starCount) {
  if (starCount === null) {
    return `<div class="rating-row"><div class="label">${escapeHtml(label)}</div><div class="small">No rating submitted</div></div>`;
  }

  const count = Math.max(1, Math.min(5, Math.round(starCount)));
  const stars = Array.from({ length: count }, (_, idx) =>
    renderStarSvg("filled", `${label}-full-${idx}`)
  ).join("");
  return `<div class="rating-row">
    <div class="label">${escapeHtml(label)}</div>
    <div class="stars-wrap">
      <div class="stars">${stars}</div>
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
      return `<div class="${className}"><span class="tag-text">${escapeHtml(
        tag.name
      )}</span>${renderTagBadges(readerLabels)}</div>`;
    }).join("")}
  </div>`;
}

function renderCompactStars(starCount) {
  if (starCount === null) {
    return `<span class="compact-stars"></span>`;
  }

  const count = Math.max(1, Math.min(5, Math.round(starCount)));
  const stars = Array.from({ length: count }, (_, idx) =>
    renderStarSvg("filled", `compact-full-${idx}`)
  ).join("");
  return `<span class="compact-stars">${stars}</span>`;
}

function renderBandRow(label, className, values) {
  const rankedValues = values
    .map((value) => BAND_ORDER.indexOf(value))
    .filter((index) => index >= 0)
    .sort((a, b) => a - b);
  const minIndex = rankedValues.length ? rankedValues[0] : null;
  const maxIndex = rankedValues.length ? rankedValues[rankedValues.length - 1] : null;

  return `<div class="band-row row-${className}">
    <div class="band-name ${className}" style="grid-column: 1; grid-row: 1;">${escapeHtml(label)}</div>
    ${
      minIndex !== null && maxIndex !== null
        ? `<div class="band-range range-${className}" style="grid-column: ${minIndex + 2} / ${maxIndex + 3}; grid-row: 1;"></div>`
        : ""
    }
    ${BAND_ORDER.map((value, idx) => {
      const inRange =
        minIndex !== null && maxIndex !== null && idx >= minIndex && idx <= maxIndex;
      return `<div class="band-chip${inRange ? " in-range" : ""}" style="grid-column: ${idx + 2}; grid-row: 1;">${escapeHtml(value)}</div>`;
    }).join("")}
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
      return `<div class="${className}"><span class="tag-text">${escapeHtml(
        tag.name
      )}</span>${renderTagBadges(readerLabels)}</div>`;
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

function renderReaderRatingRow(label, ratingValue) {
  const display = ratingValue ? String(ratingValue) : "—";
  const pillClass = ratingValue ? "reader-rating-pill" : "reader-rating-pill empty";
  return `<div class="reader-rating-row">
    <div class="reader-rating-label">${escapeHtml(label)}</div>
    <span class="${pillClass}">${escapeHtml(display)}</span>
  </div>`;
}

function renderReaderBandRow(label, value) {
  const display = value && value !== "—" ? value : "—";
  return `<div class="reader-band-row">
    <span class="reader-band-label">${escapeHtml(label)}</span>
    <span class="reader-band-value">${escapeHtml(display)}</span>
  </div>`;
}

function renderReaderTags(tags) {
  if (!tags.length) {
    return `<div class="reader-tag-empty">—</div>`;
  }
  return tags
    .map((tag) => {
      const polarity = TAG_POLARITY_MAP.get(tag);
      const className =
        polarity === "positive"
          ? "reader-tag-pill active-positive"
          : polarity === "negative"
            ? "reader-tag-pill active-negative"
            : "reader-tag-pill inactive";
      return `<div class="${className}"><span class="reader-tag-text">${escapeHtml(
        tag
      )}</span></div>`;
    })
    .join("");
}

function renderReaderNotes(reader) {
  const notes = String(reader.notes || "").trim();
  const anythingElse = String(reader.anythingElse || "").trim();
  const combined =
    notes && anythingElse
      ? `${notes}\n\n${anythingElse}`
      : notes || anythingElse || "—";
  return `<div class="reader-notes">
    <div class="label">Notes</div>
    <div class="reader-notes-body">${escapeHtml(combined)}</div>
  </div>`;
}

function renderReaderDetailPage(report) {
  const summaryText = String(report.manual.summary || "").trim() || "—";
  const nextStepsText = String(report.manual.nextSteps || "").trim() || "—";
  return `
    <section class="page reader-page">
      <div class="reader-page-main">
        <div class="reader-sections">
          ${report.readers
            .map(
              (reader) => `
            <article class="reader-section">
              <div class="reader-section-title">${escapeHtml(reader.label)}</div>
              <div class="reader-rows">
                <div class="reader-row reader-row-top">
                  <div class="reader-col ratings">
                    ${renderReaderRatingRow(
                      "I understand the candidate's reason for seeking a law degree.",
                      reader.ratingLabels.whyLaw
                    )}
                    ${renderReaderRatingRow(
                      "This candidate will thrive in law school.",
                      reader.ratingLabels.thrive
                    )}
                    ${renderReaderRatingRow(
                      "This candidate would positively contribute to a law school community.",
                      reader.ratingLabels.contribute
                    )}
                    ${renderReaderRatingRow(
                      "i feel like I know this candidate.",
                      reader.ratingLabels.know
                    )}
                  </div>
                  <div class="reader-col bands">
                    ${renderReaderBandRow("Reach", reader.bands.reach)}
                    ${renderReaderBandRow("Target", reader.bands.target)}
                    ${renderReaderBandRow("Safety", reader.bands.safety)}
                    ${renderReaderBandRow("Softs", reader.softs)}
                  </div>
                  <div class="reader-col tags">
                    ${renderReaderTags(reader.tags)}
                  </div>
                </div>
                <div class="reader-row reader-row-notes">
                  ${renderReaderNotes(reader)}
                </div>
              </div>
            </article>`
            )
            .join("")}
        </div>
      </div>
      <div class="reader-page-footer">
        <div class="reader-footer-box">
          <div class="reader-footer-title">Summary</div>
          <div class="reader-footer-body">${escapeHtml(summaryText)}</div>
        </div>
        <div class="reader-footer-box">
          <div class="reader-footer-title">Recommended Next Steps</div>
          <div class="reader-footer-body">${escapeHtml(nextStepsText)}</div>
        </div>
      </div>
    </section>
  `;
}

function renderWaitingPage() {
  return `
    <section class="page waiting-page">
      <p class="waiting-copy">waiting on design</p>
    </section>
  `;
}

function renderTagExplanationPage(tags) {
  return `
    <section class="page tag-explanation-page">
      <div class="tag-explanation-grid">
        ${tags
          .map(
            (tag) => {
              const polarity = TAG_POLARITY_MAP.get(tag);
              const className =
                polarity === "positive"
                  ? "tag-pill active-positive"
                  : polarity === "negative"
                    ? "tag-pill active-negative"
                    : "tag-pill inactive";
              return `<article class="tag-explanation-item">
              <div class="tag-explanation-title">
                <span class="${className} tag-explanation-pill">
                  <span class="tag-text">${escapeHtml(tag)}</span>
                </span>
              </div>
              <div class="tag-explanation-body">
                Lorem ipsum dolor sit amet, consectetur adipiscing elit. Sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.
              </div>
            </article>`;
            }
          )
          .join("")}
      </div>
    </section>
  `;
}

function renderTagExplanationPages(report) {
  const tags = [...(report.activeTags || new Set())].sort((a, b) =>
    a.localeCompare(b)
  );
  if (!tags.length) {
    return `
      <section class="page tag-explanation-page">
        <div class="tag-explanation-empty">No tags selected.</div>
      </section>
    `;
  }

  const pages = [];
  const perPage = 6;
  for (let i = 0; i < tags.length; i += perPage) {
    pages.push(renderTagExplanationPage(tags.slice(i, i + perPage)));
  }
  return pages.join("");
}

function getReaderSortKey(reader) {
  const profile = getReaderProfile(reader.rawReviewer);
  const name = String(profile?.fullName || reader.rawReviewer || "").trim();
  if (!name) {
    return String(reader.label || "").toLowerCase();
  }
  const parts = name.split(/\s+/);
  const key = parts.length > 1 ? parts[parts.length - 1] : parts[0];
  return key.toLowerCase();
}

function renderStudentDocument(report) {
  const summaryReaders = report.readers
    .map((reader, index) => ({ reader, index }))
    .sort((a, b) => {
      const aKey = getReaderSortKey(a.reader);
      const bKey = getReaderSortKey(b.reader);
      const byName = aKey.localeCompare(bKey);
      return byName !== 0 ? byName : a.index - b.index;
    })
    .map(({ reader }) => reader);

  return `
    <section class="page summary-page">
      <div class="summary-banner">Summary</div>
      <div class="section-block basics-section">
        <div class="section-body">
          <div class="fit-content fit-basics">
            <div class="section-title">Basics</div>
            <div class="basics-grid">
              <div class="basics-item">
                <span class="basics-label">LSAT:</span>
                <span class="basics-value">${escapeHtml(report.manual.lsat || "—")}</span>
              </div>
              <span class="basics-divider">|</span>
              <div class="basics-item">
                <span class="basics-label">GPA:</span>
                <span class="basics-value">${escapeHtml(report.manual.gpa || "—")}</span>
              </div>
              <span class="basics-divider">|</span>
              <div class="basics-item">
                <span class="basics-label">Other:</span>
                <span class="basics-value">${escapeHtml(report.manual.otherText || "—")}</span>
              </div>
            </div>
          </div>
        </div>
      </div>
      <div class="section-block readers-section">
        <div class="section-body">
          <div class="fit-content fit-readers">
          <div class="section-title">Readers</div>
          <div class="readers-card">
            <div class="reader-grid">
              ${summaryReaders
                .map((reader) => {
                  const profile = getReaderProfile(reader.rawReviewer);
                  const name = profile?.fullName || reader.rawReviewer || reader.label;
                  const avatarContent = profile?.headshotUrl
                    ? `<img src="${escapeHtml(profile.headshotUrl)}" alt="${escapeHtml(name)}" />`
                    : escapeHtml(getReaderInitials(reader.label.replace("Reader ", "R")));
                  const bio =
                    profile?.bio ||
                    "Lorem ipsum dolor sit amet, consectetur adipiscing elit. Sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat. Duis aute irure dolor in reprehenderit in voluptate velit esse cillum dolore eu fugiat nulla pariatur.";
                  return `<div class="reader-col">
                    <div class="avatar${profile?.headshotUrl ? " has-photo" : ""}">${avatarContent}</div>
                    <div>
                      <h3 class="reader-name">${escapeHtml(name)}</h3>
                      <p class="reader-bio">${escapeHtml(bio)}</p>
                    </div>
                  </div>`;
                })
                .join("")}
            </div>
          </div>
          </div>
        </div>
      </div>
      <div class="section-block key-takeaways-section">
        <div class="section-body">
          <div class="fit-content fit-takeaways">
            <div class="section-title">Key Takeaways</div>
            <div class="takeaways-card">
              <div class="takeaways-row takeaways-row-top">
                <div class="takeaway-item">
                  <div class="takeaway-title">Why Law</div>
                  ${renderCompactStars(report.summaryStars.whyLaw)}
                </div>
                <div class="takeaway-item">
                  <div class="takeaway-title">Readiness</div>
                  ${renderCompactStars(report.summaryStars.thrive)}
                </div>
                <div class="takeaway-item takeaway-softs">
                  <div class="takeaway-title">Softs</div>
                  <div class="takeaway-softs-value">${escapeHtml(report.softsDisplay || "—")}</div>
                </div>
              </div>
              <div class="takeaways-row takeaways-row-bottom">
                <div class="takeaway-item">
                  <div class="takeaway-title">Perspective</div>
                  ${renderCompactStars(report.summaryStars.contribute)}
                </div>
                <div class="takeaway-item">
                  <div class="takeaway-title">Personality</div>
                  ${renderCompactStars(report.summaryStars.know)}
                </div>
              </div>
              <div class="takeaways-bands">
                ${renderBandRow("Reach", "reach", report.bands.reach)}
                ${renderBandRow("Target", "target", report.bands.target)}
                ${renderBandRow("Safety", "safety", report.bands.safety)}
              </div>
            </div>
          </div>
        </div>
      </div>
      <div class="section-block tags-section">
        <div class="section-body">
          <div class="fit-content fit-tags">
          <div class="section-title">Tags</div>
          <div class="tags-card">
            <div class="tags-grid-wrap">
              ${renderTagGridFourColumns(report.activeTags, report.tagReaderMap)}
            </div>
          </div>
          </div>
        </div>
      </div>
    </section>

    ${renderReaderDetailPage(report)}
    ${renderTagExplanationPages(report)}
  `;
}

function getDefaultStudentInput() {
  return {
    lsat: "",
    gpa: "",
    kjd: "Not KJD",
    urm: "Non-URM",
    summary: "",
    nextSteps: "",
  };
}

function setManualControlsEnabled(enabled) {
  lsatInput.disabled = !enabled;
  gpaInput.disabled = !enabled;
  kjdSelect.disabled = !enabled;
  urmSelect.disabled = !enabled;
  summaryInput.disabled = !enabled;
  nextStepsInput.disabled = !enabled;
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
    summaryInput.value = "";
    nextStepsInput.value = "";
    inputHintEl.textContent = "";
    return;
  }

  const manual = ensureStudentInput(fileName);
  lsatInput.value = manual.lsat;
  gpaInput.value = manual.gpa;
  kjdSelect.value = manual.kjd;
  urmSelect.value = manual.urm;
  summaryInput.value = manual.summary || "";
  nextStepsInput.value = manual.nextSteps || "";
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
  fitReaderNotes(previewRoot);
  requestAnimationFrame(() => {
    fitAllSummarySections(previewRoot);
    fitReaderNotes(previewRoot);
  });
  if (document.fonts?.ready) {
    document.fonts.ready.then(() => {
      fitAllSummarySections(previewRoot);
      fitReaderNotes(previewRoot);
    });
  }
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
  manual.summary = summaryInput.value.trim();
  manual.nextSteps = nextStepsInput.value.trim();
  validateManualInputs();
}

function fitTagFonts(root = document) {
  const pills = root.querySelectorAll(".tag-pill");
  pills.forEach((pill) => {
    const textEl = pill.querySelector(".tag-text");
    if (!textEl) return;

    const prevDisplay = textEl.style.display;
    const prevClamp = textEl.style.webkitLineClamp;
    const prevOrient = textEl.style.webkitBoxOrient;

    textEl.style.display = "block";
    textEl.style.webkitLineClamp = "unset";
    textEl.style.webkitBoxOrient = "initial";

    const setFont = (sizePx) => {
      pill.style.setProperty("--tag-font-size", `${sizePx}px`);
    };

    const getLineHeight = () => {
      const computed = window.getComputedStyle(textEl);
      const fontSize = parseFloat(computed.fontSize) || TAG_FONT_BASE_PX;
      let lineHeight = parseFloat(computed.lineHeight);
      if (!lineHeight || Number.isNaN(lineHeight)) {
        lineHeight = fontSize * 1.1;
      }
      return { lineHeight, fontSize };
    };

    setFont(TAG_FONT_BASE_PX);
    let { lineHeight } = getLineHeight();
    let maxHeight = lineHeight * 2 + 0.5;
    let fullHeight = textEl.scrollHeight;

    if (fullHeight > maxHeight) {
      let low = TAG_FONT_MIN_PX;
      let high = TAG_FONT_BASE_PX;
      let best = TAG_FONT_MIN_PX;

      while (low <= high) {
        const mid = Math.floor((low + high) / 2);
        setFont(mid);
        ({ lineHeight } = getLineHeight());
        maxHeight = lineHeight * 2 + 0.5;
        fullHeight = textEl.scrollHeight;

        if (fullHeight <= maxHeight) {
          best = mid;
          low = mid + 1;
        } else {
          high = mid - 1;
        }
      }
      setFont(best);
    }

    textEl.style.display = prevDisplay;
    textEl.style.webkitLineClamp = prevClamp;
    textEl.style.webkitBoxOrient = prevOrient;
  });
}

function fitReaderNotes(root = document) {
  const notesBlocks = root.querySelectorAll(".reader-notes");
  notesBlocks.forEach((notes) => {
    const body = notes.querySelector(".reader-notes-body");
    if (!body) return;

    const label = notes.querySelector(".label");
    const computed = window.getComputedStyle(notes);
    const paddingTop = parseFloat(computed.paddingTop) || 0;
    const paddingBottom = parseFloat(computed.paddingBottom) || 0;
    const labelHeight = label ? label.offsetHeight : 0;
    const labelMarginBottom = label ? parseFloat(window.getComputedStyle(label).marginBottom) || 0 : 0;

    const availableHeight =
      notes.clientHeight - paddingTop - paddingBottom - labelHeight - labelMarginBottom;
    if (!availableHeight) return;

    const setFont = (sizePx) => {
      notes.style.setProperty("--notes-font-size", `${sizePx}px`);
    };
    const setLineHeight = (value) => {
      notes.style.setProperty("--notes-line-height", `${value}`);
    };

    setLineHeight(NOTES_LINE_BASE);
    setFont(NOTES_FONT_BASE_PX);
    if (body.scrollHeight <= availableHeight) return;

    let low = NOTES_FONT_MIN_PX;
    let high = NOTES_FONT_BASE_PX;
    let best = NOTES_FONT_MIN_PX;

    while (low <= high) {
      const mid = Math.floor((low + high) / 2);
      setFont(mid);
      if (body.scrollHeight <= availableHeight) {
        best = mid;
        low = mid + 1;
      } else {
        high = mid - 1;
      }
    }
    setFont(best);
    setLineHeight(NOTES_LINE_BASE);

    if (body.scrollHeight <= availableHeight) return;
    setLineHeight(NOTES_LINE_TIGHT);
    if (body.scrollHeight <= availableHeight) return;
    setLineHeight(NOTES_LINE_TIGHTER);
    if (body.scrollHeight <= availableHeight) return;
    setLineHeight(NOTES_LINE_MIN);
  });
}

function fitSectionContent(section) {
  const container = section.querySelector(".section-body");
  const content = container?.querySelector(".fit-content");
  if (!container || !content) return;

  content.style.transform = "scale(1)";
  content.style.width = "100%";

  const availableHeight = container.clientHeight;
  const neededHeight = content.scrollHeight;
  if (!availableHeight || !neededHeight) return;

  function applyScale(scale) {
    if (scale < 1) {
      content.style.transform = `scale(${scale})`;
      content.style.width = `${100 / scale}%`;
    } else {
      content.style.transform = "scale(1)";
      content.style.width = "100%";
    }
  }

  let scale = 1;
  const heightScale = availableHeight / neededHeight;

  // Fill vertical space first.
  if (heightScale < 1) {
    scale = heightScale;
  }

  applyScale(scale);

}

function fitAllSummarySections(root = document) {
  fitTagFonts(root);
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
    fitReaderNotes(previewRoot);
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
    summary: manual.summary,
    nextSteps: manual.nextSteps,
    otherText: `${manual.kjd} • ${manual.urm}`,
  };

  const report = buildStudentReport(selectedFile, rows, mergedManual);
  state.currentReport = report;
  state.reports = [report];
  renderPreviewFromCurrent();
  downloadPdfBtn.disabled = false;
  setStatus(`Generated report for ${selectedFile}.`);
  showValidationMessages();
}

function onStudentSelectionChange() {
  syncManualFormFromSelection();
  state.currentReport = null;
  state.reports = [];
  downloadPdfBtn.disabled = true;
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

  const hasPdfLib = Boolean(window.jspdf?.jsPDF);
  const hasHtmlToImage = Boolean(window.htmlToImage?.toPng);
  const hasHtml2Canvas = typeof window.html2canvas === "function";

  if (!hasPdfLib || (!hasHtmlToImage && !hasHtml2Canvas)) {
    state.warnings = [
      ...state.warnings,
      "PDF library is unavailable. Use your browser Print dialog as fallback.",
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
  staging.style.left = "-10000px";
  staging.style.top = "0";
  staging.style.width = "612px";
  staging.style.opacity = "1";
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
    fitReaderNotes(staging);
    await new Promise((resolve) => requestAnimationFrame(resolve));
    fitAllSummarySections(staging);
    fitReaderNotes(staging);
    await new Promise((resolve) => requestAnimationFrame(resolve));

    const filename = `${sanitizeFileName(report.fileName) || "student-report"}.pdf`;
    const pageNodes = [...clone.querySelectorAll(".page")];
    if (!pageNodes.length) {
      throw new Error("No pages found to export.");
    }
    const pdf = new window.jspdf.jsPDF({
      unit: "pt",
      format: [612, 792],
      orientation: "portrait",
    });

    for (let i = 0; i < pageNodes.length; i += 1) {
      let imageData = "";
      if (hasHtmlToImage) {
        imageData = await window.htmlToImage.toPng(pageNodes[i], {
          pixelRatio: 2,
          cacheBust: true,
          backgroundColor: "#ffffff",
          canvasWidth: 612,
          canvasHeight: 792,
          width: 612,
          height: 792,
        });
      } else {
        const canvas = await window.html2canvas(pageNodes[i], {
          scale: 2,
          useCORS: true,
          backgroundColor: "#ffffff",
          width: 612,
          height: 792,
          windowWidth: 612,
          windowHeight: 792,
          x: 0,
          y: 0,
          scrollX: 0,
          scrollY: 0,
        });
        imageData = canvas.toDataURL("image/jpeg", 0.98);
      }

      if (i > 0) {
        pdf.addPage([612, 792], "portrait");
      }
      pdf.addImage(imageData, "PNG", 0, 0, 612, 792, undefined, "FAST");
    }

    pdf.save(filename);
  } catch (_error) {
    state.warnings = [
      ...state.warnings,
      "PDF export failed. Use your browser Print dialog as fallback.",
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
