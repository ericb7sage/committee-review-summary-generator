const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');
const path = require('node:path');
const root = path.join(__dirname, '..');
// Run the browser's recommendation logic without starting its UI or network loads.
const source = fs.readFileSync(path.join(root, 'web/app.js'), 'utf8')
  .replace(/initializeAppVersion\(\);\nvoid ensureLogoDataUri\(\);\nvoid Promise.all\(\[initializeTagCatalog\(\), initializeSchoolCatalog\(\)\]\);/, '');
const context = vm.createContext({
  document: { getElementById: () => ({ addEventListener() {} }) },
  window: { addEventListener() {}, location: { href: 'http://localhost/' } },
  URL,
  console,
});
vm.runInContext(source, context);
context.catalog = JSON.parse(fs.readFileSync(path.join(root, 'web/school-rankings.json'), 'utf8'));
vm.runInContext('setSchoolCatalog(catalog)', context);
const run = (code) => vm.runInContext(code, context);
run(`
  var row = { Reviewer: 'Reader' };
  Object.values(SCHOOL_RECOMMENDATION_FIELDS_TWO).forEach(fields => {
    row[fields[0]] = 'University of Toronto';
    row[fields[1]] = '';
  });
`);
assert.equal(run(`validateStudentSchoolRecommendations('Student', [row]).length`), 0);
run(`var recommendations = resolveReaderSchoolRecommendations(row, 'schools-two');`);
assert.equal(run('recommendations.reach[0].rank'), null);
assert.match(run(`renderReaderSchoolRow('Reach', 'reach', recommendations.reach)`), /University of Toronto.*Unranked/s);
const allUnranked = run(`renderSchoolRecommendationPlot([{schoolRecommendations: recommendations}], '2026')`);
assert.match(allUnranked, /University of Toronto/);
assert.match(allUnranked, /No ranked schools to plot/);
assert.doesNotMatch(allUnranked, /school-rank-dot|school-range-span|NaN|Infinity/);
run(`recommendations.target.push(resolveSchoolRecommendation(catalog.schools[0].name));`);
const mixed = run(`renderSchoolRecommendationPlot([{schoolRecommendations: recommendations}], '2026')`);
assert.match(mixed, /University of Toronto/);
assert.equal((mixed.match(/class="school-rank-dot /g) || []).length, 1);
assert.match(mixed, /school-range-span target[^>]+--range-lane:1/);
assert.doesNotMatch(mixed, /school-range-span (reach|safety)|NaN|Infinity/);
run(`row[SCHOOL_RECOMMENDATION_FIELDS_TWO.reach[1]] = 'university of toronto';`);
assert.match(run(`validateStudentSchoolRecommendations('Student', [row]).join(' ')`), /twice/);
run(`row[SCHOOL_RECOMMENDATION_FIELDS_TWO.reach[0]] = '';`);
assert.match(run(`validateStudentSchoolRecommendations('Student', [row]).join(' ')`), /missing required/);
assert.equal(run(`validateSchoolCatalog({cycle: '2026', schools: [{name: 'Toronto', rank: null}]}).length`), 0);
assert.ok(run(`validateSchoolCatalog({cycle: '2026', schools: [{name: 'Invalid', rank: ''}]}).length`) > 0);
assert.equal(fs.readFileSync(path.join(root, 'docs/app.js'), 'utf8'), fs.readFileSync(path.join(root, 'web/app.js'), 'utf8'));
console.log('Unranked school regression checks passed.');
