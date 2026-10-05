# Generic Internship Report Template

This project generates a reusable internship-report DOCX template for students of the **Islamic University of Technology**, **Department of Computer Science and Engineering**, enrolled in the **BSc in Computer Science and Engineering** program.

The template is suitable for internships in any industry, organization, role or technology area. It does not contain details from a particular previous internship or project. The generated report is saved as:

```text
industrial-training-report.docx
```

## Requirements

- Node.js 18 or newer
- npm
- The `docx` package installed from `package-lock.json`

Node.js 22 is also supported and was used to validate the current generator.

## Install dependencies

From this directory, run:

```bash
npm install
```

## Generate the report

The recommended command is:

```bash
npm run generate
```

This runs the template `generate.js` builder and overwrites the existing `industrial-training-report.docx`.

You can also run the generator directly:

```bash
node generate.js
```

`generate.js` is the shareable template builder. It contains no narrative from a previous internship or project. The university, department and degree are fixed to the IUT CSE program. Edit the remaining fields in the `INFO` object near the top of the file, replace the highlighted content prompts with the trainee's actual information, and then run the command. Values left as `[FILL: ...]` remain yellow-highlighted in the generated document.

## Fields to complete

Update these fields in `generate.js`:

- Student name
- Student ID
- Organization name
- Internship or training program
- Group or cohort
- Internship start and end dates
- Internship mode
- Internship venue
- Industry supervisor
- Main project name

The following fields are already configured and should normally not be changed:

- Islamic University of Technology
- Department of Computer Science and Engineering
- BSc in Computer Science and Engineering

## Optional logo

If `iut_logo.png` is present in this directory, it is used on the cover page. The generator continues to work if the logo is unavailable.

## Output validation

After generation, validate that the DOCX archive is not corrupted:

```bash
unzip -t industrial-training-report.docx
```

The command should end with:

```text
No errors detected in compressed data
```

## Complete the template

The personalized report source is in `build_report.js`. The template source is in `generate.js`. To create a completed report from the template:

```bash
1. Open `generate.js`.
2. Edit the fillable fields in the `INFO` object near the top.
3. Replace the highlighted prompts throughout the report with the trainee's actual organization profile, activities, responsibilities, project details, tools, results, challenges, learning outcomes and references.
4. Remove or replace every remaining `[FILL: ...]` item before final submission.
5. Run `npm run generate`.
```

The output file is regenerated from the JavaScript source; editing the DOCX directly will not update the generator.

## Main files

| File | Purpose |
|---|---|
| `generate.js` | Generic IUT CSE internship-template builder |
| `industrial-training-report.docx` | Generated Word document |
| `iut_logo.png` | Optional cover-page logo |
| `package.json` | npm command and dependency declaration |
| `package-lock.json` | Locked dependency versions |
