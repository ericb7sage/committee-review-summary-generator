import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const sourcePath = path.join(root, "data", "7sage-rankings-2026.json");
const outputPaths = [
  path.join(root, "web", "school-rankings.json"),
  path.join(root, "docs", "school-rankings.json"),
];

const aliasOverrides = {
  UChicago: ["University of Chicago", "University of Chicago Law School"],
  UPenn: ["University of Pennsylvania", "Penn", "Penn Carey Law", "University of Pennsylvania Carey Law School"],
  UVA: ["University of Virginia", "University of Virginia School of Law"],
  NYU: ["New York University", "New York University School of Law"],
  UCLA: ["University of California Los Angeles", "University of California—Los Angeles", "UCLA School of Law"],
  WashU: ["Washington University", "Washington University in St. Louis", "Washington University School of Law"],
  Berkeley: ["UC Berkeley", "University of California Berkeley", "University of California—Berkeley", "Berkeley Law"],
  "UT Austin": ["University of Texas", "University of Texas at Austin", "Texas Law"],
  UNC: ["University of North Carolina", "University of North Carolina at Chapel Hill", "UNC School of Law"],
  BYU: ["Brigham Young University", "J. Reuben Clark Law School"],
  USC: ["University of Southern California", "USC Gould", "USC Gould School of Law"],
  FSU: ["Florida State University", "Florida State University College of Law"],
  "UC - Irvine": ["UC Irvine", "University of California Irvine", "University of California—Irvine"],
  "U Florida": ["University of Florida", "University of Florida Levin College of Law", "UF Law"],
  "Illinois - Urbana": ["University of Illinois", "University of Illinois Urbana-Champaign", "University of Illinois College of Law"],
  "Indiana Bloomington": ["Indiana University Bloomington", "Indiana University Maurer School of Law", "Maurer"],
  "UC - Davis": ["UC Davis", "University of California Davis", "University of California—Davis", "UC Davis School of Law"],
  "U Washington": ["University of Washington", "University of Washington School of Law"],
  "Colorado - Boulder": ["University of Colorado", "University of Colorado Boulder", "Colorado Law"],
  USD: ["University of San Diego", "University of San Diego School of Law"],
  UCONN: ["University of Connecticut", "University of Connecticut School of Law", "UConn Law"],
  Cardozo: ["Benjamin N. Cardozo School of Law", "Yeshiva University Cardozo School of Law"],
  "Penn State": ["Penn State Dickinson Law", "Pennsylvania State University Dickinson Law"],
  "Pennsylvania State - Penn State Law": ["Penn State Law", "Pennsylvania State University Penn State Law"],
  "Loyola Marymount - LA": ["Loyola Marymount University", "Loyola Law School Los Angeles", "LMU Loyola Law School"],
  FIU: ["Florida International University", "Florida International University College of Law"],
  GSU: ["Georgia State University", "Georgia State University College of Law"],
  "Loyola - Chicago": ["Loyola Chicago", "Loyola University Chicago", "Loyola University—Chicago"],
  LSU: ["Louisiana State University", "LSU Paul M. Hebert Law Center"],
  "UC Law San Francisco": ["University of California College of the Law San Francisco", "UC Hastings", "UC Hastings College of the Law"],
  "Nevada - Las Vegas": ["University of Nevada Las Vegas", "UNLV", "UNLV Boyd School of Law"],
  SLU: ["Saint Louis University", "Saint Louis University School of Law"],
  "Case Western": ["Case Western Reserve University", "Case Western Reserve University School of Law"],
  "Missouri - Kansas City": ["University of Missouri Kansas City", "UMKC", "UMKC School of Law"],
  "U Arkansas - Fayetteville": ["University of Arkansas", "University of Arkansas School of Law", "Arkansas Fayetteville"],
  "Chicago-Kent": ["Chicago-Kent College of Law", "Illinois Institute of Technology Chicago-Kent College of Law"],
  "St. Thomas (Minnesota)": ["University of St. Thomas Minnesota", "University of St. Thomas School of Law Minnesota"],
  American: ["American University", "American University Washington College of Law", "Washington College of Law"],
  "Lewis And Clark": ["Lewis & Clark", "Lewis & Clark Law School", "Lewis and Clark College"],
  UNM: ["University of New Mexico", "University of New Mexico School of Law"],
  "IU McKinney": ["Indiana University McKinney", "Indiana University Robert H. McKinney School of Law"],
  WVU: ["West Virginia University", "West Virginia University College of Law"],
  "Loyola - New Orleans": ["Loyola University New Orleans", "Loyola University New Orleans College of Law"],
  "U Arkansas - Little Rock": ["University of Arkansas at Little Rock", "UA Little Rock William H. Bowen School of Law", "Bowen"],
  "Pacific (Mcgeorge)": ["University of the Pacific", "McGeorge School of Law", "University of the Pacific McGeorge School of Law"],
  UND: ["University of North Dakota", "University of North Dakota School of Law"],
  USF: ["University of San Francisco", "University of San Francisco School of Law"],
  UNTD: ["University of North Texas at Dallas", "UNT Dallas College of Law"],
  NIU: ["Northern Illinois University", "Northern Illinois University College of Law"],
  "Illinois - Chicago": ["University of Illinois Chicago", "UIC Law", "UIC School of Law"],
  "New England - Boston": ["New England Law Boston", "New England Law | Boston"],
  CUNY: ["City University of New York", "CUNY School of Law"],
  "UMass - Dartmouth": ["University of Massachusetts Dartmouth", "UMass Law", "UMass School of Law"],
  "Florida A&M": ["Florida Agricultural and Mechanical University", "Florida A&M University College of Law", "FAMU"],
  NCCU: ["North Carolina Central University", "North Carolina Central University School of Law"],
  ONU: ["Ohio Northern University", "Ohio Northern University Pettit College of Law"],
  "SIU - Carbondale": ["Southern Illinois University", "Southern Illinois University School of Law", "SIU Law"],
  UDC: ["University of the District of Columbia", "UDC David A. Clarke School of Law"],
  "Western Michigan": ["Western Michigan University", "Western Michigan University Cooley Law School", "Cooley"],
  "Southern Methodist": ["Southern Methodist University", "SMU", "SMU Dedman School of Law"],
  Denver: ["University of Denver", "University of Denver Sturm College of Law", "Denver Sturm"],
};

function normalize(value) {
  return String(value || "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/[’‘]/g, "'")
    .replace(/[‐‑‒–—]/g, "-")
    .replace(/&/g, " and ")
    .replace(/[^a-zA-Z0-9]+/g, " ")
    .replace(/\s+/g, " ")
    .trim()
    .toLowerCase();
}

function generatedAliases(name) {
  const aliases = [`${name} Law School`, `${name} School of Law`];
  if (/^[A-Za-z.'-]+$/.test(name)) {
    aliases.push(`${name} University`, `University of ${name}`);
  }
  return aliases;
}

const sourceRows = JSON.parse(fs.readFileSync(sourcePath, "utf8"));
const schools = sourceRows.map(({ name, rank }) => ({
  name: String(name).replace(/\s+/g, " ").trim(),
  rank: rank === "175+" ? 175 : Number(rank),
  ...(rank === "175+" ? { rankLabel: "175+" } : {}),
  aliases: [],
}));

const canonicalOwners = new Map(schools.map((school, index) => [normalize(school.name), index]));
const aliasOwners = new Map(canonicalOwners);

schools.forEach((school, index) => {
  const candidates = [
    ...generatedAliases(school.name),
    ...(aliasOverrides[school.name] || []),
  ];
  const aliases = [];
  for (const alias of candidates) {
    const trimmed = String(alias).replace(/\s+/g, " ").trim();
    const key = normalize(trimmed);
    if (!key || key === normalize(school.name) || aliases.some((item) => normalize(item) === key)) continue;
    const owner = aliasOwners.get(key);
    if (owner !== undefined && owner !== index) continue;
    aliasOwners.set(key, index);
    aliases.push(trimmed);
  }
  school.aliases = aliases;
});

const catalog = {
  cycle: "2026 U.S. News rankings",
  source: "https://7sage.com/admissions/rankings",
  retrieved: "2026-08-04",
  maxRank: 175,
  schools,
};

const serialized = `${JSON.stringify(catalog, null, 2)}\n`;
for (const outputPath of outputPaths) fs.writeFileSync(outputPath, serialized);
console.log(`Wrote ${schools.length} schools to ${outputPaths.length} catalogs.`);
