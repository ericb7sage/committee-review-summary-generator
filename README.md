# Committee Review Summary Generator (Frontend Only)

This is a browser-only JavaScript app for your Coda CSV to printable PDF workflow.

## What it does right now

- Loads a Coda-generated CSV directly in the browser
- Enforces a fixed CSV header contract
- Groups rows by `File` and generates one report per student
- Requires exactly 3 reviews per student to generate a report
- Converts rubric text to 1-5 scores and averages each category
- Renders rounded SVG stars for the averaged categories
- Displays all 40 tags with conditional activation styling
- Places each selected tag's explanation after the first reader ballot that selected it, without repeating shared explanations
- Displays readers from fewest to most selected visible tags, with reader number as the tie-breaker
- Starts every new reader ballot on a fresh report page
- Lets users enter per-student metadata (LSAT, GPA, KJD, URM)
- Generates an on-screen report preview before enabling PDF download
- Lets users switch between students in a multi-student CSV and review each report before downloading
- Keeps Print Current Student as a fallback

## Quick start

Open `web/index.html` in a browser.

## GitHub Pages

Set Pages source to `main` + `/docs`. `/docs` is a published copy of `/web`.

## Centralized tag explanations (no server)

Tag explanations are centralized in:

- `docs/tags.json` (source of truth for GitHub Pages)

The app loads this file at runtime. If it cannot be loaded or fails validation, the app falls back to built-in defaults and shows a warning.

### Coworker update workflow

1. Open `docs/tags.json` in GitHub
2. Click edit (pencil icon)
3. Update tag `description` values (and optionally `hidden`)
4. Commit to `main`
5. Wait for GitHub Pages to republish
6. Hard refresh the app (`Cmd/Ctrl + Shift + R`)

### `tags.json` format

Each item must include:

- `name` (string, unique)
- `polarity` (`positive` or `negative`)
- `description` (string)

Optional:

- `hidden` (boolean; if true, hidden from the page-1 tag grid)

## Current flow

1. Upload CSV
2. Select a student (`File`)
3. Select a document title (`Application Autopsy` or `Committee Review`)
4. Enter LSAT, GPA, KJD status, and URM status
5. Click **Generate Report**
6. Click **Download PDF** (or use **Print Current Student** as fallback)

## Aggregation behavior

- Grouping key: exact `File` value
- Only students with exactly 3 review rows are rendered
- Reviewers are anonymized as `Reader 1`, `Reader 2`, `Reader 3`
- Reader order is randomized every generation run
- Page 4 shows individual reader comments from:
  - `Notes`
  - `Anything Else?`

## Required CSV headers

- `Reviewer`
- `File`
- `Why Law?`
- `Thrive?`
- `Contribute?`
- `Know?`
- `Tags`
- `Reach`
- `Target`
- `Safety`
- `Notes`
- `Anything Else?`

## School recommendation plot

Exports can use two school recommendations per category with these six columns:

- `Recommend a Reach 1`
- `Recommend a Reach 2`
- `Recommend a Target 1`
- `Recommend a Target 2`
- `Recommend a Safety 1`
- `Recommend a Safety 2`

The first recommendation in each category is required for every reader; the second is optional. The summary reserves six compact positions per category (two for each of three readers), leaving an intentional blank when an optional school is omitted. Reader ballots use the same two-position layout. Repeated recommendations across readers remain in the lists and collapse to a counted consensus marker on the number line.

The existing one-school format remains supported:

- `Recommend a Reach`
- `Recommend a Target`
- `Recommend a Safety`

Coda exports using `Recommend Two Reaches` and `Recommend Two Targets` are normalized to the corresponding one-school fields. Despite those column labels, a single populated school remains valid and the second summary position stays blank.

Each of the three review rows must name one school in every required recommendation column. School names are matched against the versioned `school-rankings.json` catalog; exported rank columns are intentionally ignored. Unknown schools and duplicate choices within the same reader/category block that student's report. If the recommendation columns are absent, the report continues to use the legacy `Reach`, `Target`, and `Safety` tier plot.

When nearby school ranks are staggered above or below the number line, a category-colored connector marks the bubble's exact position on the shared baseline.

The catalog contains the 2026 U.S. News rankings published on the 7Sage rankings page. Its source snapshot lives in `data/7sage-rankings-2026.json`; run `node scripts/build-school-catalog.mjs` after editing the source or alias overrides to regenerate identical `web` and `docs` catalogs. Schools in the published `175+` group use rank 175 for plotting while retaining the `175+` display label.

## Template key fields

The app maps values from these columns to display text:

- `Why Law?`
- `Thrive?`
- `Contribute?`
- `Know?`
- `Reach`
- `Target`
- `Safety`

Unknown keys are flagged in validation warnings and rendered as `Unknown key: ...`.

## Report layout (4 pages)

1. Aggregate rubric ratings (stars + numeric)
2. Reach / Target / Safety summary
3. 40-tag activation grid (positive/negative styling)
4. Reader comments (anonymized)
