# mathtype-docx

通过 MathType COM 引擎，将 LaTeX 公式插入 Word .docx 文件，生成原生 OMML 方程对象。

## 原理

```
LaTeX 输入  →  Equation.DSMT4 (MathType COM)  →  MathML  →  OMML  →  python-docx  →  .docx
```

MathType 的 `Equation.DSMT4` COM 对象可**独立运行**，无需打开 Word。转换步骤完全脱离 Word COM，仅最终 XML 插入依赖 `python-docx`。

## 快速开始

```bash
# 插入单个公式
python scripts/docx_equation.py \
  --docx paper.docx \
  --latex "r_{i,t}=100\times(\ln P_{i,t}-\ln P_{i,t-1})" \
  --pos "after:研究方法"

# 纯 LaTeX → OMML 转换（不动 docx）
python scripts/latex_to_omml.py --latex "\hat{\mu}_{i,e}=\frac{1}{T_0}\sum r_{i,t}"

# 批量升级文档中所有文本公式为 OMML 方程
python scripts/docx_equation.py --docx paper.docx --batch
```

## 环境要求

- **Windows**（MathType 仅支持 Windows）
- **MathType 7+** 已安装（在 7.4.4 上测试通过）
- Python 3.10+，安装依赖：

```bash
pip install python-docx pywin32 lxml
```

## 支持的 LaTeX 命令

| 类别 | 命令 |
|------|------|
| 上下标 | `x_{i}`, `x^{2}` |
| 分式 | `\frac{a}{b}` |
| 根号 | `\sqrt{x}`, `\sqrt[n]{x}` |
| 希腊字母 | `\alpha`, `\beta`, `\mu`, `\sigma`, `\Sigma`, `\Delta` |
| 运算符 | `\times`, `\cdot`, `\pm`, `\div`, `\sum`, `\prod`, `\int` |
| 关系符 | `\leq`, `\geq`, `\neq`, `\approx`, `\equiv` |
| 重音符 | `\hat{x}`, `\bar{x}`, `\tilde{x}` |
| 括号 | `\left(`, `\right)`, `\left[`, `\right]`, `\left\{`, `\right\}` |
| 函数 | `\ln`, `\log`, `\sin`, `\cos`, `\tan`, `\exp`, `\max`, `\min` |
| 特殊符号 | `\infty`, `\partial`, `\nabla`, `\forall`, `\exists`, `\in`, `\subset` |

## 命令行参考

### `docx_equation.py`

```
--docx PATH      目标 .docx 文件路径
--latex STRING   要插入的 LaTeX 公式
--pos WHERE      插入位置："end"（末尾）、"after:文本"、输入段落序号
--replace N      将第 N 段替换为 OMML 方程
--batch          批量转换文档中所有文本公式
--output PATH    输出路径（默认覆盖原文件）
```

### `latex_to_omml.py`

```
--latex STRING   单条 LaTeX 公式
--file PATH      包含 LaTeX 公式的文本文件（每行一条）
--output PATH    OMML XML 输出路径
```

## 架构说明

转换管道分两阶段：

1. **LaTeX → MathML**（`Equation.DSMT4` COM）— MathType 渲染引擎解析 LaTeX，输出标准 MathML 2.0
2. **MathML → OMML**（`latex_to_omml.py:mathml_to_omml()`）— 纯 Python 递归转换器，将 MathML 元素映射为 Office Math Markup Language

OMML 方程通过 python-docx 的 XML 操作直接注入 `.docx` ZIP 包，无需 Word COM。

## 典型应用

- **学术论文**：在方法论章节插入事件研究法标准公式（AR、CAR、t 统计量等）
- **毕业论文**：LaTeX 公式一键转为 Word 原生方程，导师可直接双击编辑
- **批量排版**：将整篇文档中手工输入的文本公式统一升级为 OMML 方程对象

## 局限性

- 依赖 MathType 7+ 已安装且 COM 注册正常
- MathML→OMML 转换器覆盖常用 presentation-MathML 子集，极端边缘情况可能需要手动调整
- 公式默认以块级居中插入（如需行内公式需要额外设置）

## 许可证

MIT
