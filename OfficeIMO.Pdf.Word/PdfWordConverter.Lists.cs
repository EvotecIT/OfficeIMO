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
            private readonly PdfCore.PdfLogicalListMarker?[] _markers = new PdfCore.PdfLogicalListMarker?[MaximumLevels];
            private readonly int[] _starts = new int[MaximumLevels];
            private readonly int?[] _counters = new int?[MaximumLevels];
            private readonly bool[] _carried = new bool[MaximumLevels];
            private WordList? _list;

            internal SourceListNumbering(WordDocument document) {
                _document = document;
            }

            internal void End() => _list = null;

            internal WordParagraph AddItem(PdfCore.PdfLogicalListItem item) {
                int level = Math.Max(0, Math.Min(MaximumLevels - 1, item.Level - 1));
                PdfCore.PdfLogicalListMarker marker = item.MarkerInfo;
                if (_list == null) Start(carriedLevels: 0);
                else if (!Continues(level, marker)) Start(carriedLevels: level);
                WordList list = _list!;
                if (_markers[level] == null) {
                    SeedParents(list, level, marker);
                    ConfigureLevel(list, level, marker, Math.Max(0, marker.Ordinal ?? 1));
                }

                _counters[level] = marker.Ordinal ?? (_counters[level] + 1 ?? _starts[level]);
                for (int deeper = level + 1; deeper < MaximumLevels; deeper++) _counters[deeper] = null;
                return list.AddItem((string?)null, level);
            }

            private bool Continues(int level, PdfCore.PdfLogicalListMarker marker) {
                // A parent level carried into a restarted list only shows its number; its next item needs a list
                // that counts on from there.
                if (_carried[level]) return false;
                if (_markers[level] is not PdfCore.PdfLogicalListMarker current) return true;
                if (current.Format != marker.Format || current.Prefix != marker.Prefix || current.Suffix != marker.Suffix ||
                    current.IncludesParentNumbers != marker.IncludesParentNumbers) return false;
                if (!marker.Ordinal.HasValue) return true;
                // Word numbers the next item from the counter; a marker that does not follow is a new list.
                int expected = _counters[level] + 1 ?? _starts[level];
                return marker.Ordinal.Value == expected;
            }

            /// <summary>
            /// Starts a new numbering instance. A nested list that restarts keeps its parent levels, at their current
            /// numbers, so "2.5" after "2." still shows the 2.
            /// </summary>
            private void Start(int carriedLevels) {
                var parents = new (PdfCore.PdfLogicalListMarker Marker, int Number)?[carriedLevels];
                for (int level = 0; level < carriedLevels; level++) {
                    if (_markers[level] is PdfCore.PdfLogicalListMarker marker) parents[level] = (marker, _counters[level] ?? _starts[level]);
                }

                WordList list = _document.AddCustomList();
                for (int level = 0; level < MaximumLevels; level++) {
                    WordListLevel listLevel = list.Numbering.AddLevel(new WordListLevel(WordListLevelKind.Bullet));
                    listLevel.IndentationLeft = OfficeLength.Points(36D + level * 18D);
                    listLevel.IndentationHanging = OfficeLength.Points(18D);
                    _markers[level] = null;
                    _starts[level] = 1;
                    _counters[level] = null;
                    _carried[level] = false;
                }
                _list = list;

                for (int level = 0; level < carriedLevels; level++) {
                    if (parents[level] is not { } parent) continue;
                    ConfigureLevel(list, level, parent.Marker, parent.Number);
                    _counters[level] = parent.Number;
                    _carried[level] = true;
                }
            }

            /// <summary>An outline number with no parent item before it ("2.5" first) sets up the parents it shows.</summary>
            private void SeedParents(WordList list, int level, PdfCore.PdfLogicalListMarker marker) {
                if (!marker.IncludesParentNumbers || marker.ParentNumbers.Count != level) return;
                for (int parent = 0; parent < level; parent++) {
                    if (_markers[parent] != null) continue;
                    int number = marker.ParentNumbers[parent];
                    ConfigureLevel(list, parent, PdfCore.PdfLogicalListMarker.Parse(number.ToString(CultureInfo.InvariantCulture) + "."), number);
                    _counters[parent] = number;
                    _carried[parent] = true;
                }
            }

            private void ConfigureLevel(WordList list, int level, PdfCore.PdfLogicalListMarker marker, int start) {
                _markers[level] = marker;
                if (!marker.IsOrdered) return;
                WordListLevel listLevel = list.Numbering.Levels[level];
                listLevel.NumberFormat = marker.Format switch {
                    PdfCore.PdfLogicalListNumberFormat.LowerLetter => WordNumberFormat.LowerLetter,
                    PdfCore.PdfLogicalListNumberFormat.UpperLetter => WordNumberFormat.UpperLetter,
                    PdfCore.PdfLogicalListNumberFormat.LowerRoman => WordNumberFormat.LowerRoman,
                    PdfCore.PdfLogicalListNumberFormat.UpperRoman => WordNumberFormat.UpperRoman,
                    _ => WordNumberFormat.Decimal
                };
                string suffix = marker.Suffix.Length == 0 && marker.Prefix.Length == 0 ? "." : marker.Suffix;
                var levelText = new StringBuilder(marker.Prefix);
                // "1.2" repeats its parent numbers, which Word writes as "%1.%2".
                for (int shown = marker.IncludesParentNumbers ? 0 : level; shown <= level; shown++) {
                    levelText.Append('%').Append((shown + 1).ToString(CultureInfo.InvariantCulture));
                    if (shown < level) levelText.Append('.');
                }
                listLevel.LevelText = levelText.Append(suffix).ToString();
                _starts[level] = start;
                listLevel.StartNumberingValue = start;
            }
        }
    }
}
