# MathType DOCX Equation Skill

Insert and edit professional MathType/LaTeX equations in Word .docx files using MathType's OLE engine for LaTeX-to-MathML conversion, then embedding as native OMML equations.

## What This Skill Does

- Converts LaTeX equations to native Word OMML equations (via MathType COM bridge)
- Inserts equations into existing .docx files at specified locations
- Replaces plain-text equations or Unicode math with proper OMML/equation objects
- Batch-converts all formulas in a document

## When To Use

- User says "插入公式", "add equation", "LaTeX to Word", "docx equation", "mathtype formula"
- User needs professional math notation in a .docx
- User wants to convert LaTeX formulas to Word equations
- User wants to upgrade plain-text formulas in a document to proper equation objects

## Architecture

```
LaTeX input → Equation.DSMT4 (MathType COM) → MathML → OMML converter → python-docx → .docx with equations
```

The MathType COM engine (`Equation.DSMT4`) works **standalone** — no Word COM required for conversion. Only python-docx is needed for document insertion.

## Prerequisites

- MathType 7+ installed (tested with 7.4.4)
- Python: `python-docx`, `pywin32`, `lxml`
- Windows OS (MathType is Windows-only)

## Usage

### 1. Insert a single equation into a DOCX

```bash
python scripts/docx_equation.py --docx "paper.docx" --latex "r_{i,t}=100\times(\ln P_{i,t}-\ln P_{i,t-1})" --position "after:（三）事件研究方法"
```

Options:
- `--latex`: LaTeX formula string
- `--position`: Where to insert — "after:TEXT" (after paragraph containing TEXT), "before:TEXT", "end", or paragraph index
- `--eqnum`: Equation number (optional)
- `--display`: "block" or "inline" (default: block)

### 2. Batch-convert all formulas in a document

```bash
python scripts/docx_equation.py --docx "paper.docx" --batch
```

This finds all paragraphs formatted as equations (centered, italic formulas in the document) and converts them to proper OMML equations.

### 3. Convert LaTeX string to OMML (without DOCX)

```bash
python scripts/latex_to_omml.py --latex "E=mc^2"
```

Returns the OMML XML string.

## Supported LaTeX Commands

| Category | Commands |
|----------|----------|
| Subscripts/Superscripts | `x_{i}`, `x^{2}`, `x_{i}^{2}` |
| Fractions | `\frac{a}{b}` |
| Square roots | `\sqrt{x}`, `\sqrt[n]{x}` |
| Greek letters | `\alpha`, `\beta`, `\gamma`, `\delta`, `\mu`, `\sigma`, `\Sigma`, `\Delta` |
| Operators | `\times`, `\cdot`, `\pm`, `\mp`, `\div` |
| Relations | `\leq`, `\geq`, `\neq`, `\approx`, `\equiv` |
| Sums/Products | `\sum`, `\prod`, `\int` |
| Accents | `\hat{x}`, `\bar{x}`, `\tilde{x}`, `\vec{x}` |
| Arrows | `\rightarrow`, `\leftarrow`, `\Rightarrow` |
| Brackets | `\left(`, `\right)`, `\left[`, `\right]`, `\left\{`, `\right\}` |
| Text/Functions | `\text{...}`, `\ln`, `\log`, `\sin`, `\cos`, `\tan`, `\exp`, `\max`, `\min` |
| Special | `\infty`, `\partial`, `\nabla`, `\forall`, `\exists`, `\in`, `\notin`, `\subset` |

## Implementation Notes

- MathType's `Equation.DSMT4` COM object is used for LaTeX→MathML conversion (standalone, no Word needed)
- MathML output is parsed with lxml and converted to OMML XML
- OMML equations are inserted into the DOCX via python-docx `_element` XML manipulation
- Fallback to Unicode math text if MathType COM is unavailable
- Equation objects render as native Word equations (editable in Word's equation editor)

## Known Limitations

- MathType COM must be installed and registered
- The MathML→OMML converter handles the common MathML subset; edge cases may need manual adjustment
- OMML equations are inserted as block-level elements; inline equations require additional formatting
