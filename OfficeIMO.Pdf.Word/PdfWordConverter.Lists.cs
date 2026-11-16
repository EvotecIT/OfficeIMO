using System.Globalization;
using System.Text;
using OfficeIMO.Word;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Pdf.Word {
    internal static partial class PdfWordConverter {
        /// <summary>
        /// Writes PDF list items as Word numbering. Each source list gets its own numbering instance, so a second list
        /// starts again instead of counting on from the first. Each level takes its format (bullet, decimal, letter or
        /// Roman, with the source punctuation) and start value from the first item drawn at that level.
        /// </summary>
        private sealed class SourceListNumbering {
            private const int MaximumLevels = 9;
            private readonly WordDocument _document;
            private readonly PdfCore.PdfLogicalListNumberFormat?[] _formats = new PdfCore.PdfLogicalListNumberFormat?[MaximumLevels];
            private readonly int[] _starts = new int[MaximumLevels];
            private readonly int?[] _counters = new int?[MaximumLevels];
            private WordList? _list;

            internal SourceListNumbering(WordDocument document) {
                _document = document;
            }

            internal void End() => _list = null;

            internal WordParagraph AddItem(PdfCore.PdfLogicalListItem item) {
                int level = Math.Max(0, Math.Min(MaximumLevels - 1, item.Level - 1));
                PdfCore.PdfLogicalListMarker marker = item.MarkerInfo;
                if (_list == null || !Continues(level, marker)) Start();
                WordList list = _list!;
                if (_formats[level] == null) ConfigureLevel(list, level, marker);

                _counters[level] = marker.Ordinal ?? (_counters[level] + 1 ?? _starts[level]);
                for (int deeper = level + 1; deeper < MaximumLevels; deeper++) _counters[deeper] = null;
                return list.AddItem((string?)null, level);
            }

            private bool Continues(int level, PdfCore.PdfLogicalListMarker marker) {
                PdfCore.PdfLogicalListNumberFormat? format = _formats[level];
                if (format == null) return true;
                if (format.Value != marker.Format) return false;
                if (!marker.Ordinal.HasValue) return true;
                // Word numbers the next item from the counter; a marker that does not follow is a new list.
                int expected = _counters[level] + 1 ?? _starts[level];
                return marker.Ordinal.Value == expected;
            }

            private void Start() {
                WordList list = _document.AddCustomList();
                for (int level = 0; level < MaximumLevels; level++) {
                    WordListLevel listLevel = list.Numbering.AddLevel(new WordListLevel(WordListLevelKind.Bullet));
                    listLevel.IndentationLeft = OfficeLength.Points(36D + level * 18D);
                    listLevel.IndentationHanging = OfficeLength.Points(18D);
                    _formats[level] = null;
                    _starts[level] = 1;
                    _counters[level] = null;
                }
                _list = list;
            }

            private void ConfigureLevel(WordList list, int level, PdfCore.PdfLogicalListMarker marker) {
                _formats[level] = marker.Format;
                if (!marker.IsOrdered) return;
                WordListLevel listLevel = list.Numbering.Levels[level];
                listLevel.NumberFormat = marker.Format switch {
                    PdfCore.PdfLogicalListNumberFormat.LowerLetter => WordNumberFormat.LowerLetter,
                    PdfCore.PdfLogicalListNumberFormat.UpperLetter => WordNumberFormat.UpperLetter,
                    PdfCore.PdfLogicalListNumberFormat.LowerRoman => WordNumberFormat.LowerRoman,
                    PdfCore.PdfLogicalListNumberFormat.UpperRoman => WordNumberFormat.UpperRoman,
                    _ => WordNumberFormat.Decimal
                };
                string suffix = marker.Suffix.Length == 0 && marker.Prefix.Length == 0 && !marker.IncludesParentNumbers ? "." : marker.Suffix;
                var levelText = new StringBuilder(marker.Prefix);
                // "1.2" repeats its parent numbers, which Word writes as "%1.%2".
                for (int shown = marker.IncludesParentNumbers ? 0 : level; shown <= level; shown++) {
                    levelText.Append('%').Append((shown + 1).ToString(CultureInfo.InvariantCulture));
                    if (shown < level) levelText.Append('.');
                }
                listLevel.LevelText = levelText.Append(suffix).ToString();
                _starts[level] = Math.Max(0, marker.Ordinal ?? 1);
                listLevel.StartNumberingValue = _starts[level];
            }
        }
    }
}
