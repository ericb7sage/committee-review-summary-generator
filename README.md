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
- Lets users enter per-student metadata (LSAT, GPA, KJD, URM)
- Supports one-click PDF download for the currently generated student
- Keeps Print Current Student as a fallback

## Quick start

Open `web/index.html` in a browser.

## GitHub Pages

Set Pages source to `main` + `/docs`. `/docs` is a published copy of `/web`.

## Current flow

1. Upload CSV
2. Select a student (`File`)
3. Enter LSAT, GPA, KJD status, and URM status
4. Click **Generate Report**
5. Click **Download PDF** (or use **Print Current Student** as fallback)

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
