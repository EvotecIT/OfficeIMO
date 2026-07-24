using OfficeIMO.PowerPoint.LegacyPpt.Diagnostics;
using OfficeIMO.PowerPoint.LegacyPpt.Internal;
using OfficeIMO.PowerPoint.LegacyPpt.Model;

namespace OfficeIMO.PowerPoint.LegacyPpt {
    public sealed partial class LegacyPptPresentation {
        private void ParseFontCollection(LegacyPptRecord document, LegacyPptImportOptions options) {
            LegacyPptRecord? collection = document.DescendantsAndSelf().FirstOrDefault(record =>
                record.Type == RecordFontCollection);
            if (collection == null) return;
            LegacyPptRecord[] children = collection.Children.ToArray();
            for (int childIndex = 0; childIndex < children.Length; childIndex++) {
                LegacyPptRecord atom = children[childIndex];
                if (atom.Type != RecordFontEntityAtom) continue;
                if (atom.PayloadLength != 68) {
                    AddDiagnostic("PPT-FONT-ENTITY-TRUNCATED", LegacyPptDiagnosticSeverity.Warning,
                        "A FontEntityAtom does not have the required 68-byte payload and was skipped.",
                        atom.Offset);
                    continue;
                }
                ushort index = atom.Instance;
                if (_fontsByIndex.ContainsKey(index)) {
                    AddDiagnostic("PPT-FONT-INDEX-DUPLICATE", LegacyPptDiagnosticSeverity.Warning,
                        $"Font index {index} is defined more than once; the first definition is retained.",
                        atom.Offset);
                    continue;
                }
                string typeface = atom.ReadUtf16Text(0, 64);
                int terminator = typeface.IndexOf('\0');
                if (terminator >= 0) typeface = typeface.Substring(0, terminator);
                typeface = LegacyPptXmlText.SanitizeAttributeValue(typeface) ?? string.Empty;
                if (typeface.Length == 0) {
                    AddDiagnostic("PPT-FONT-NAME-EMPTY", LegacyPptDiagnosticSeverity.Warning,
                        $"Font index {index} has an empty typeface name and was skipped.", atom.Offset);
                    continue;
                }
                bool hasEmbeddedData = false;
                for (int next = childIndex + 1; next < children.Length
                        && children[next].Type != RecordFontEntityAtom; next++) {
                    hasEmbeddedData |= children[next].Type == RecordFontEmbedDataBlob;
                }
                byte embedFlags = atom.ReadByte(65);
                byte fontTypeFlags = atom.ReadByte(66);
                var font = new LegacyPptFont(index, typeface, atom.ReadByte(64),
                    isEmbeddedSubset: (embedFlags & 0x01) != 0,
                    isRaster: (fontTypeFlags & 0x01) != 0,
                    isDevice: (fontTypeFlags & 0x02) != 0,
                    isTrueType: (fontTypeFlags & 0x04) != 0,
                    disableSubstitution: (fontTypeFlags & 0x08) != 0,
                    pitchAndFamily: atom.ReadByte(67), hasEmbeddedData);
                _fonts.Add(font);
                _fontsByIndex.Add(index, font);
                if (hasEmbeddedData && options.ReportUnsupportedContent) {
                    AddDiagnostic("PPT-FONT-EMBEDDED-PRESERVE-ONLY", LegacyPptDiagnosticSeverity.Warning,
                        $"Embedded data for typeface '{typeface}' is preserved but not projected to the Open XML package.",
                        atom.Offset);
                }
            }
            _fonts.Sort((left, right) => left.Index.CompareTo(right.Index));
        }
    }
}
