# mathtype-docx

Insert LaTeX-rendered equations into Word .docx files as native OMML equations — via MathType's COM engine.

## How It Works

```
LaTeX input  →  Equation.DSMT4 (MathType COM)  →  MathML  →  OMML  →  python-docx  →  .docx
```

MathType's `Equation.DSMT4` COM object runs **standalone** — no Word COM required for the conversion step. Only `python-docx` is needed for the final XML insertion.

## Install as Claude Code Skill

```bash
# 1. Clone the repo
git clone git@github.com:zb1655235-cpu/mathtype-docx.git

# 2. Install to Claude Code skills directory
#    Windows (this repo's author uses this path):
mkdir -p "$HOME/claude-config/skills"
cp -r mathtype-docx "$HOME/claude-config/skills/"

#    macOS/Linux:
mkdir -p ~/.claude/skills
cp -r mathtype-docx ~/.claude/skills/

# 3. Install Python dependencies
pip install python-docx pywin32 lxml
```

After installing, restart Claude Code. The skill activates automatically when you say things like "插入公式", "add equation to docx", or "convert LaTeX to Word equation".

> **Note:** Claude Code looks for skills in `~/.claude/skills/` by default. If you use a custom config path, adjust step 2 accordingly. The skill directory must contain `SKILL.md` at its root.

## Quick Start

```bash
# Insert a single equation
python scripts/docx_equation.py \
  --docx paper.docx \
  --latex "r_{i,t}=100\times(\ln P_{i,t}-\ln P_{i,t-1})" \
  --pos "after:Methodology"

# Convert LaTeX to OMML (without touching a docx)
python scripts/latex_to_omml.py --latex "\hat{\mu}_{i,e}=\frac{1}{T_0}\sum r_{i,t}"

# Upgrade all plain-text formulas in a document to OMML equations
python scripts/docx_equation.py --docx paper.docx --batch
```

## Requirements

- **Windows** (MathType is Windows-only)
- **MathType 7+** installed (tested with 7.4.4)
- Python 3.10+ with `python-docx`, `pywin32`, `lxml`

```bash
pip install python-docx pywin32 lxml
```

## Supported LaTeX

| Category | Commands |
|----------|----------|
| Sub/Superscripts | `x_{i}`, `x^{2}` |
| Fractions | `\frac{a}{b}` |
| Square roots | `\sqrt{x}`, `\sqrt[n]{x}` |
| Greek | `\alpha`, `\beta`, `\mu`, `\sigma`, `\Sigma`, `\Delta` |
| Operators | `\times`, `\cdot`, `\pm`, `\div`, `\sum`, `\prod`, `\int` |
| Relations | `\leq`, `\geq`, `\neq`, `\approx`, `\equiv` |
| Accents | `\hat{x}`, `\bar{x}`, `\tilde{x}` |
| Brackets | `\left(`, `\right)`, `\left[`, `\right]`, `\left\{`, `\right\}` |
| Functions | `\ln`, `\log`, `\sin`, `\cos`, `\tan`, `\exp`, `\max`, `\min` |
| Special | `\infty`, `\partial`, `\nabla`, `\forall`, `\exists`, `\in`, `\subset` |

## CLI Reference

### `docx_equation.py`

```
--docx PATH      Path to .docx file
--latex STRING   LaTeX formula to insert
--pos WHERE      Insert position: "end", "after:TEXT", "before:TEXT", or paragraph index
--replace N      Replace paragraph at index N with equation
--batch          Batch-convert all text formulas in the document to OMML
--output PATH    Output path (default: overwrite input)
```

### `latex_to_omml.py`

```
--latex STRING   Single LaTeX formula
--file PATH      File with one LaTeX formula per line
--output PATH    Output file for OMML XML
```

## Architecture

The conversion pipeline has two stages:

1. **LaTeX → MathML** (`Equation.DSMT4` COM) — MathType's rendering engine parses LaTeX and emits standard MathML 2.0
2. **MathML → OMML** (`latex_to_omml.py:mathml_to_omml()`) — A pure-Python recursive transformer mapping MathML elements to Office Math Markup Language

OMML equations are then injected into the `.docx` ZIP via python-docx's XML manipulation (no Word COM needed for insertion).

## Limitations

- Requires MathType 7+ installed and COM-registered on the machine
- The MathML→OMML converter handles the common presentation-MathML subset; edge cases may need tuning
- Equations are inserted as block-level (centered) by default

## License

MIT
