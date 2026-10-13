# Cyber Persistent Engagement

Python tools for dictionary-based coding of cybersecurity strategy documents that have already been converted into structured Excel workbooks.

The scripts classify words and phrases into **deterrence**, **norms**, and **persistent engagement**, flag ambiguous matches, and aggregate scores for comparative research.

## Research context

This repository supports my cybersecurity-document research work at the University of Cincinnati. The broader work involves multilingual documents and cross-national comparison; these scripts cover workbook scoring and aggregation. Document extraction and translation are separate steps.

Related research reference: [DOI: 10.7945/n1h7-wv34](https://doi.org/10.7945/n1h7-wv34).

## Programs

| Script | Role | Dependencies |
| --- | --- | --- |
| [FormulaScore.py](FormulaScore.py) | Codes and sorts entries, adds highlights, and writes Excel formulas for scores. | openpyxl, pandas |
| [StrategyScore.py](StrategyScore.py) | Earlier scoring implementation that preserves entry order and writes numeric scores. | openpyxl |
| [MastersheetGenerator.py](MastersheetGenerator.py) | Reads coded workbooks through Excel and collects results in a master workbook. | openpyxl, xlwings, desktop Excel |

## Setup

Use Python 3. For the master-workbook step, use a desktop environment with Excel available to xlwings.

```bash
git clone https://github.com/stevenrashid/Cyber-Peristent-Engagement.git
cd Cyber-Peristent-Engagement
python -m pip install openpyxl pandas xlwings
```

The input workbooks and dictionary are **not included**. Supply compatible files before running the scripts.

### Workbook format

The scripts use the active worksheet and fixed cell positions:

| Section | Text | Frequency | Code | Comment | Total / relative scores |
| --- | --- | --- | --- | --- | --- |
| Phrases | A | B | C | D | G / H |
| Words | J | K | L | M | P / Q |
| Words and phrases | S | T | U | V | Y / Z |

Data begins in row 2. Each text section must be contiguous because scanning stops at its first blank cell. Frequencies must be numeric.

The dictionary's active worksheet stores deterrence terms in column A, norms in C, and persistent-engagement terms in E, starting at row 2. Dictionary terms are uppercased when read; prepare input text consistently because the matching functions compare it as stored.

## Run the workflow

1. Create `Templated/`, `Coded/`, and `Dictionaries/` in the repository root.
2. Place formatted input workbooks in `Templated/` and the dictionary at `Dictionaries/Final_Dictionary.xlsx`, or edit the configuration variables at the top of `FormulaScore.py`.
3. Run:
   ```bash
   python FormulaScore.py
   ```
4. Review the generated `new_*.xlsx` files in `Coded/`, including flagged matches.
5. With Excel available, run:
   ```bash
   python MastersheetGenerator.py
   ```

The result is `Mastersheet.xlsx`, with `PHRASES`, `WORDS`, and `WORDS&PHRASES` sheets. Keep only intended `.xlsx` inputs in the coded folder. The master generator extracts country letters and year digits from filenames such as `new_Canada2020.xlsx`.

Use `StrategyScore.py` as an alternative to `FormulaScore.py`, not as an additional scoring stage. Its default paths are `Input Folder`, `Output Folder`, and `Dictionary.xlsx`; configure them before execution.

## Codes and review flags

| Code | Category | Highlight |
| --- | --- | --- |
| 0 | No dictionary match | None |
| 1 | Deterrence | Green |
| 2 | Norms | Blue |
| 3 | Persistent engagement | Purple |

Orange comments indicate partial matches; red comments indicate matches across categories. Conflicting matches require review because the coding cell retains the last assigned category.

## Operational notes

- Use a fresh output directory or archive previous outputs before rerunning; existing `new_` files can cause move conflicts.
- FormulaScore sorts entries by code, which changes their original order.
- Formula-based scores require spreadsheet recalculation; openpyxl alone does not calculate them.
- Empty or zero-frequency sections can produce division-by-zero results. Validate inputs and review workbook outputs.
- Dictionary matching is a research coding aid and depends on the dictionary, formatting, and human review.
