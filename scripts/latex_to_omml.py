#!/usr/bin/env python3
r"""
LaTeX to OMML converter using MathType's Equation.DSMT4 COM engine.

Architecture:
  LaTeX --[Equation.DSMT4.SetLaTeX()]--> MathML --[mathml_to_omml()]--> OMML

Usage:
  python latex_to_omml.py --latex "r_{i,t}=100\times(\ln P_{i,t}-\ln P_{i,t-1})"
  python latex_to_omml.py --file formulas.txt
"""

import argparse
import sys
import io
from lxml import etree

# OMML namespace
OMML_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/math'
MATHML_NS = 'http://www.w3.org/1998/Math/MathML'

OMML_MAP = {
    'm': OMML_NS,
    'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
}


def _q(n):
    """Create qualified OMML name."""
    return f'{{{OMML_NS}}}{n}'


def _make_run(text, style='i'):
    """Create m:r element with m:t child.
    style: 'i' (italic/default variable), 'n' (number), 't' (text/normal)
    """
    r = etree.Element(_q('r'))
    if style == 't':
        nor = etree.SubElement(r, _q('nor'))
        etree.SubElement(nor, _q('t')).text = text
    else:
        t = etree.SubElement(r, _q('t'))
        t.text = text
        if style == 'i':
            t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
    return r


def _mathml_elem_to_omml(elem):
    """Convert a MathML element to OMML elements (list)."""
    tag = elem.tag.replace(f'{{{MATHML_NS}}}', '')
    result = []

    if tag == 'math':
        for child in elem:
            result.extend(_mathml_elem_to_omml(child))
        return result

    elif tag == 'mrow':
        for child in elem:
            result.extend(_mathml_elem_to_omml(child))
        return result

    elif tag == 'mi':
        text = ''.join(elem.itertext())
        result.append(_make_run(text, 'i'))
        return result

    elif tag == 'mn':
        text = ''.join(elem.itertext())
        result.append(_make_run(text, 'n'))
        return result

    elif tag == 'mo':
        text = ''.join(elem.itertext())
        # Handle common operator conversions
        text = text.replace('×', '×')  # × stays
        text = text.replace('−', '−')  # − stays
        result.append(_make_run(text, 'i'))
        return result

    elif tag == 'mtext':
        text = ''.join(elem.itertext())
        result.append(_make_run(text, 't'))
        return result

    elif tag == 'msub':
        sub = etree.Element(_q('sSub'))
        e = etree.SubElement(sub, _q('e'))
        base_children = []
        sub_children = []
        found_base = False
        for child in elem:
            child_tag = child.tag.replace(f'{{{MATHML_NS}}}', '')
            if child_tag == 'mrow':
                if not found_base:
                    base_children.extend(_mathml_elem_to_omml(child))
                    found_base = True
                else:
                    sub_children.extend(_mathml_elem_to_omml(child))
            elif not found_base:
                base_children.extend(_mathml_elem_to_omml(child))
                found_base = True
            else:
                sub_children.extend(_mathml_elem_to_omml(child))

        for bc in base_children:
            e.append(bc)
        sub_e = etree.SubElement(sub, _q('sub'))
        for sc in sub_children:
            sub_e.append(sc)
        result.append(sub)
        return result

    elif tag == 'msup':
        sup = etree.Element(_q('sSup'))
        e = etree.SubElement(sup, _q('e'))
        sup_children = []
        found_base = False
        for child in elem:
            child_tag = child.tag.replace(f'{{{MATHML_NS}}}', '')
            if child_tag == 'mrow':
                if not found_base:
                    for c in _mathml_elem_to_omml(child):
                        e.append(c)
                    found_base = True
                else:
                    sup_children.extend(_mathml_elem_to_omml(child))
            elif not found_base:
                for c in _mathml_elem_to_omml(child):
                    e.append(c)
                found_base = True
            else:
                sup_children.extend(_mathml_elem_to_omml(child))

        sup_e = etree.SubElement(sup, _q('sup'))
        for sc in sup_children:
            sup_e.append(sc)
        result.append(sup)
        return result

    elif tag == 'msubsup':
        ss = etree.Element(_q('sSubSup'))
        e = etree.SubElement(ss, _q('e'))
        sub = etree.SubElement(ss, _q('sub'))
        sup = etree.SubElement(ss, _q('sup'))
        # Simple heuristic: first child=base, second=sub, third=sup
        children = list(elem)
        if len(children) >= 1:
            for c in _mathml_elem_to_omml(children[0]):
                e.append(c)
        if len(children) >= 2:
            for c in _mathml_elem_to_omml(children[1]):
                sub.append(c)
        if len(children) >= 3:
            for c in _mathml_elem_to_omml(children[2]):
                sup.append(c)
        result.append(ss)
        return result

    elif tag == 'mfrac':
        frac = etree.Element(_q('f'))
        num = etree.SubElement(frac, _q('num'))
        den = etree.SubElement(frac, _q('den'))
        found_num = False
        for child in elem:
            for c in _mathml_elem_to_omml(child):
                if not found_num:
                    num.append(c)
                else:
                    den.append(c)
            found_num = True
        result.append(frac)
        return result

    elif tag == 'msqrt':
        rad = etree.Element(_q('rad'))
        e = etree.SubElement(rad, _q('e'))
        for child in elem:
            for c in _mathml_elem_to_omml(child):
                e.append(c)
        result.append(rad)
        return result

    elif tag == 'mroot':
        rad = etree.Element(_q('rad'))
        e = etree.SubElement(rad, _q('e'))
        deg = etree.SubElement(rad, _q('deg'))
        children = list(elem)
        if len(children) >= 1:
            for c in _mathml_elem_to_omml(children[0]):
                e.append(c)
        if len(children) >= 2:
            for c in _mathml_elem_to_omml(children[1]):
                deg.append(c)
        result.append(rad)
        return result

    elif tag == 'mover':
        acc = etree.Element(_q('acc'))
        e = etree.SubElement(acc, _q('e'))
        accent = etree.SubElement(acc, _q('accPr'))
        # Simplified accent handling
        children = list(elem)
        if len(children) >= 1:
            for c in _mathml_elem_to_omml(children[0]):
                e.append(c)
        if len(children) >= 2:
            acc_text = ''.join(children[1].itertext())
            if acc_text == '̂' or acc_text == '^':
                acc.set(_q('accChr'), '̂')
            elif acc_text == '¯' or acc_text == '̄':
                acc.set(_q('accChr'), '̄')
            elif acc_text == '̃' or acc_text == '~':
                acc.set(_q('accChr'), '̃')
        result.append(acc)
        return result

    elif tag == 'munder':
        # Underscript (e.g., limits)
        group = etree.Element(_q('limLow'))
        e = etree.SubElement(group, _q('e'))
        lim = etree.SubElement(group, _q('lim'))
        children = list(elem)
        if len(children) >= 1:
            for c in _mathml_elem_to_omml(children[0]):
                e.append(c)
        if len(children) >= 2:
            for c in _mathml_elem_to_omml(children[1]):
                lim.append(c)
        result.append(group)
        return result

    elif tag == 'munderover':
        # Summation/product with limits
        nary = etree.Element(_q('nary'))
        e = etree.SubElement(nary, _q('e'))
        sub = etree.SubElement(nary, _q('sub'))
        sup = etree.SubElement(nary, _q('sup'))

        children_text = [''.join(c.itertext()) for c in elem]

        # Detect nary operator
        if any('∑' in ct for ct in children_text):
            nary.set(_q('chr'), '∑')  # Σ
        elif any('∏' in ct for ct in children_text):
            nary.set(_q('chr'), '∏')  # ∏
        elif any('∫' in ct for ct in children_text):
            nary.set(_q('chr'), '∫')  # ∫

        for i, child in enumerate(elem):
            child_tag = child.tag.replace(f'{{{MATHML_NS}}}', '')
            if i == 0:
                for c in _mathml_elem_to_omml(child):
                    e.append(c)
            elif i == 1:
                for c in _mathml_elem_to_omml(child):
                    sub.append(c)
            elif i == 2:
                for c in _mathml_elem_to_omml(child):
                    sup.append(c)
        result.append(nary)
        return result

    elif tag == 'mfenced':
        # Parentheses/brackets
        left_char = elem.get('open', '(')
        right_char = elem.get('close', ')')

        delim = etree.Element(_q('d'))
        dPr = etree.SubElement(delim, _q('dPr'))
        beg_chr = etree.SubElement(dPr, _q('begChr'))
        beg_chr.set(_q('val'), left_char)
        end_chr = etree.SubElement(dPr, _q('endChr'))
        end_chr.set(_q('val'), right_char)
        e = etree.SubElement(delim, _q('e'))

        for child in elem:
            for c in _mathml_elem_to_omml(child):
                e.append(c)
        result.append(delim)
        return result

    elif tag == 'mstyle':
        for child in elem:
            result.extend(_mathml_elem_to_omml(child))
        return result

    elif tag == 'mphantom':
        for child in elem:
            result.extend(_mathml_elem_to_omml(child))
        return result

    elif tag == 'mspace':
        return result  # Skip spacing elements

    elif tag == 'merror':
        text = ''.join(elem.itertext())
        print(f"  [WARN] MathML error element: {text}", file=sys.stderr)
        result.append(_make_run(f'[ERR:{text}]', 't'))
        return result

    elif tag == 'semantics':
        # Take the first child (usually presentation MathML)
        for child in elem:
            if child.tag.replace(f'{{{MATHML_NS}}}', '') in ('mrow', 'mi', 'mo', 'msub', 'msup', 'mfrac', 'msqrt'):
                result.extend(_mathml_elem_to_omml(child))
                return result
        return result

    else:
        # Unknown element: extract text and wrap
        text = ''.join(elem.itertext())
        if text.strip():
            print(f"  [WARN] Unknown MathML element <{tag}>: '{text[:50]}'", file=sys.stderr)
            result.append(_make_run(text, 'i'))
        return result


def mathml_to_omml(mathml_str):
    """Convert MathML string to OMML XML element."""
    # Parse MathML
    mathml_str = mathml_str.strip()
    if not mathml_str:
        raise ValueError("Empty MathML string")

    # Handle HTML entity &#xXXXX; → unicode characters
    import html
    mathml_str = html.unescape(mathml_str)

    root = etree.fromstring(mathml_str.encode('utf-8'))

    # Create OMML oMath element
    omath = etree.Element(_q('oMath'))

    # Convert children
    for child in root:
        omml_children = _mathml_elem_to_omml(child)
        for oc in omml_children:
            omath.append(oc)

    return omath


def latex_to_omml(latex_formula, use_mathtype=True):
    """
    Convert LaTeX formula to OMML XML.

    Args:
        latex_formula: LaTeX formula string
        use_mathtype: Use MathType COM for conversion (default True).
                      Falls back to text if COM unavailable.

    Returns:
        (omml_element, source) tuple where source is 'mathtype' or 'text_fallback'
    """
    if use_mathtype:
        try:
            import win32com.client
            eq = win32com.client.Dispatch("Equation.DSMT4")
            eq.SetLaTeX(latex_formula)
            mathml = eq.GetMathML
            omml = mathml_to_omml(mathml)
            return omml, 'mathtype'
        except ImportError:
            print("  [INFO] pywin32 not available, using text fallback", file=sys.stderr)
        except Exception as e:
            print(f"  [WARN] MathType COM error: {e}, using text fallback", file=sys.stderr)

    # Text fallback: wrap in a simple OMML run
    omath = etree.Element(_q('oMath'))
    r = etree.SubElement(omath, _q('r'))
    t = etree.SubElement(r, _q('t'))
    t.text = latex_formula
    t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
    return omath, 'text_fallback'


def main():
    parser = argparse.ArgumentParser(description='Convert LaTeX to OMML via MathType')
    parser.add_argument('--latex', type=str, help='LaTeX formula string')
    parser.add_argument('--file', type=str, help='File with one LaTeX formula per line')
    parser.add_argument('--output', type=str, help='Output file for OMML XML')

    args = parser.parse_args()

    formulas = []
    if args.latex:
        formulas = [args.latex]
    elif args.file:
        with open(args.file, 'r', encoding='utf-8') as f:
            formulas = [line.strip() for line in f if line.strip()]

    for i, formula in enumerate(formulas):
        omml, source = latex_to_omml(formula)
        xml_str = etree.tostring(omml, encoding='unicode', pretty_print=True)
        print(f"Formula {i+1} [{source}]:")
        print(xml_str)
        print()

    return 0


if __name__ == '__main__':
    sys.exit(main())
