"""Regenerate the built-in OfficeTheme values from Build/Visual/themes.json (Python 3)."""
import json
import re
from pathlib import Path

root = Path(__file__).resolve().parent
source = root.parents[1] / 'Build' / 'Visual' / 'themes.json'
themes = json.loads(source.read_text(encoding='utf-8'))['themes']
ROLES = {'Accent1', 'Accent2', 'Accent3', 'Accent4', 'Accent5', 'Accent6', 'Text', 'TextMuted', 'Background',
         'Success', 'Warning', 'Danger'}
DOCS = {
    'Modern': 'a clean sans-serif theme with a cobalt accent and airy, tinted-header tables (the default)',
    'Classic': 'a traditional serif document theme with navy headings and an ochre title rule',
    'Minimal': 'a restrained theme without fills whose fonts have metric-compatible open substitutes, so layout matches across office suites',
    'Boardroom': 'a formal theme with regular-weight serif headings, a gold rule and generous margins',
    'Dark': 'a screen theme with light text on a dark background; Word and Excel keep a light page',
}


def color(value):
    if value in ROLES:
        return 'OfficeColor.' + value
    if not re.fullmatch(r'[0-9A-F]{6}', value):
        raise ValueError('Unexpected colour: ' + value)
    return 'C("' + value + '")'


def num(value):
    text = repr(float(value))
    return text[:-2] if text.endswith('.0') else text


def pt(value):
    return 'P(' + num(value) + ')'


def font(entry):
    return 'F("' + entry['family'] + '", ' + ', '.join('"' + f + '"' for f in entry['fallbacks']) + ')'


def rule(value, default_width=None):
    if value is None:
        return 'OfficeThemeRule.None'
    if isinstance(value, str):
        # Footer rules name a colour only; the design draws them as a hairline.
        return 'new OfficeThemeRule(' + color(value) + ', ' + pt(default_width) + ')'
    return 'new OfficeThemeRule(' + color(value['color']) + ', ' + pt(value['width']) + ')'


def spacing(pair):
    return 'S(' + num(pair[0]) + ', ' + num(pair[1]) + ')'


MARKERS = {'none': 'None', 'circle': 'Circle'}
BORDERS = {'horizontal': 'Horizontal', 'grid': 'Grid', 'none': 'None'}
ALIGN = {'left': 'Left', 'right': 'Right', 'center': 'Center'}

lines = ['// Generated from Build/Visual/themes.json by generate_theme_builtins.py. Do not edit by hand.',
         '#nullable enable', 'using OfficeIMO.Drawing;', '', 'namespace OfficeIMO;', '',
         'public sealed partial class OfficeTheme {']
for t in themes:
    lines += ['    /// <summary>Gets ' + DOCS[t['name']] + '.</summary>',
              '    public static OfficeTheme ' + t['name'] + ' { get; } = Create' + t['name'] + '();']
for t in themes:
    p, f, ts, tb, ch, sp, ex, de = (t[k] for k in ('palette', 'fonts', 'typeScale', 'tableStyle', 'chartStyle', 'spacing', 'extensions', 'derived'))
    tx = ex['table']
    cx = ex['chart']
    pg = ex['page']
    ft = ex['footer']
    hs = ex['headingSpacing']
    lines += ['', '    private static OfficeTheme Create' + t['name'] + '() {',
              '        var palette = new OfficeThemePalette(' + ', '.join(color(p[k]) for k in (
                  'background', 'text', 'textMuted', 'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6',
                  'success', 'warning', 'danger')) + ');',
              '        var fonts = new OfficeThemeFonts(' + font(f['heading']) + ', ' + font(f['body']) + ');',
              '        var typeScale = new OfficeThemeTypeScale(' + ', '.join(pt(ts[k]) for k in (
                  'title', 'heading1', 'heading2', 'heading3', 'heading4', 'heading5', 'heading6', 'body', 'caption')) + ');',
              '        var table = new OfficeThemeTableStyle(' + color(tb['headerFill']) + ', ' + color(tb['headerText']) + ', '
              + str(tb['banding']).lower() + ', ' + color(tb['bandFill']) + ', ' + color(tb['borderColor']) + ', '
              + pt(tb['borderWidth']) + ', ' + pt(tb['cellPadding']) + ',',
              '            OfficeThemeTableBorderMode.' + BORDERS[tx['borderMode']] + ', ' + rule(tx['headerRule']) + ', '
              + str(tx['headerBold']).lower() + ', OfficeAlignment.' + ALIGN[tx['numbersAlign']] + ');',
              '        var chart = new OfficeThemeChartStyle(new[] { ' + ', '.join(color(c) for c in ch['seriesPalette']) + ' },',
              '            ' + font(ch['axisFont']) + ', ' + font(ch['titleFont']) + ',',
              '            ' + pt(ch['axisFontSize']) + ', ' + pt(ch['titleFontSize']) + ', ' + color(ch['gridlineColor']) + ', '
              + str(ch['showDataLabels']).lower() + ', ' + str(ch['showLegend']).lower() + ',',
              '            ' + color(cx['axisLine']) + ', ' + color(cx['axisText']) + ', ' + pt(cx['gridlineWidth']) + ', '
              + pt(cx['lineWidth']) + ', ' + str(int(cx['barGapPercent'])) + ', OfficeChartMarkerShape.' + MARKERS[cx['markers']] + ');',
              '        var spacing = new OfficeThemeSpacing(' + ', '.join(pt(sp[k]) for k in (
                  'paragraphBefore', 'paragraphAfter', 'listIndent', 'slideMarginTop', 'slideMarginRight', 'slideMarginBottom',
                  'slideMarginLeft')) + ');',
              '        var text = new OfficeThemeTextStyle(' + num(ex['lineSpacing']) + ', ' + color(ex['titleColor']) + ', '
              + color(ex['subtitleColor']) + ', ' + color(ex['headingColor']) + ', ' + color(ex['link']) + ', '
              + str(ex['headingBold']).lower() + ', ' + rule(ex.get('titleRule')) + ',',
              '            ' + ', '.join(spacing(hs[k]) for k in ('title', 'heading1', 'heading2', 'heading3')) + ');',
              '        var surfaces = new OfficeThemeSurfaces(' + ', '.join(color(v) for v in (
                  ex['surface'], ex['rule'], de['accent1Tint'], de['successFill'], de['warningFill'], de['dangerFill'])) + ');',
              '        var page = new OfficeThemePage(' + ', '.join(pt(pg[k]) for k in ('marginTop', 'marginRight', 'marginBottom', 'marginLeft')) + ');',
              '        var footer = new OfficeThemeFooter(' + pt(ft['size']) + ', ' + color(ft['color']) + ', ' + rule(ft['rule'], 0.5) + ');',
              '        return new OfficeTheme("' + t['name'] + '", palette, fonts, typeScale, table, chart, spacing, text, surfaces, page, footer);',
              '    }']
lines += ['',
          '    private static OfficeColor C(string value) => OfficeColor.ParseHex(value);',
          '    private static OfficeLength P(double value) => OfficeLength.Points(value);',
          '    private static OfficeThemeFont F(string family, params string[] fallbacks) => new OfficeThemeFont(family, fallbacks);',
          '    private static OfficeThemeParagraphSpacing S(double before, double after) => new OfficeThemeParagraphSpacing(P(before), P(after));',
          '}', '']
(root / 'OfficeTheme.BuiltIns.Generated.cs').write_text('\n'.join(lines), encoding='utf-8')
