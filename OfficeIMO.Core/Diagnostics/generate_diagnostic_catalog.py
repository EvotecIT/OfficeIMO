"""Regenerate the shared C# catalog from diagnostic-code-map.json (Python 3)."""
import json
from pathlib import Path

root = Path(__file__).resolve().parent
catalog = json.loads((root / 'diagnostic-code-map.json').read_text(encoding='utf-8'))
groups = {}
for entry in catalog['entries']:
    code, member = entry['catalogCode'], entry['member']
    groups.setdefault(member, code)
    if groups[member] != code:
        raise ValueError('Conflicting catalog member: ' + member)
groups.setdefault('ChartFlattened', 'chart-flattened')
for first, last in [('A', 'F'), ('G', 'P'), ('Q', 'Z')]:
    lines = ['// Generated from diagnostic-code-map.json by generate_diagnostic_catalog.py.',
             'namespace OfficeIMO;', '',
             '/// <summary>Stable diagnostic codes shared by document formats and converters.</summary>',
             'public static partial class OfficeDiagnosticCodes {']
    for member, code in sorted(groups.items()):
        if first <= member[0] <= last:
            lines += ['    /// <summary>Stable code: <c>' + code + '</c>.</summary>',
                      '    public const string ' + member + ' = "' + code + '";']
    lines += ['}', '']
    (root / ('OfficeDiagnosticCodes.' + first + last + '.Generated.cs')).write_text('\n'.join(lines), encoding='utf-8')
