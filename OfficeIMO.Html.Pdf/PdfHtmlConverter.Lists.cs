using System.Globalization;
using System.Text;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Html.Pdf;

public static partial class PdfHtmlConverterExtensions {
    /// <summary>
    /// Nests consecutive PDF list items into &lt;ul&gt; and &lt;ol&gt; lists. A deeper item opens a list inside the
    /// previous entry; a marker of another kind, or a number that does not follow on, starts a new list.
    /// </summary>
    private sealed class HtmlListWriter {
        private readonly StringBuilder _builder;
        private readonly List<OpenList> _open = new();

        internal HtmlListWriter(StringBuilder builder) {
            _builder = builder;
        }

        internal void AppendItem(PdfCore.PdfLogicalListItem item, string content) {
            int level = Math.Max(1, item.Level);
            PdfCore.PdfLogicalListMarker marker = item.MarkerInfo;
            while (_open.Count > 0 && _open[_open.Count - 1].Level > level) CloseLast();

            if (_open.Count > 0 && _open[_open.Count - 1].Level == level) {
                OpenList current = _open[_open.Count - 1];
                if (current.Format == marker.Format && (!marker.Ordinal.HasValue || marker.Ordinal.Value == current.Next)) {
                    _builder.Append("</li>");
                } else {
                    CloseLast();
                    Open(level, marker);
                }
            } else {
                Open(level, marker);
            }

            OpenList list = _open[_open.Count - 1];
            list.Next = (marker.Ordinal ?? list.Next) + 1;
            _builder.Append("<li>").Append(content);
        }

        internal void Close() {
            while (_open.Count > 0) CloseLast();
        }

        private void Open(int level, PdfCore.PdfLogicalListMarker marker) {
            bool ordered = marker.IsOrdered;
            _builder.Append(ordered ? "<ol" : "<ul");
            string? type = marker.Format switch {
                PdfCore.PdfLogicalListNumberFormat.LowerLetter => "a",
                PdfCore.PdfLogicalListNumberFormat.UpperLetter => "A",
                PdfCore.PdfLogicalListNumberFormat.LowerRoman => "i",
                PdfCore.PdfLogicalListNumberFormat.UpperRoman => "I",
                _ => null
            };
            if (type != null) _builder.Append(" type=\"").Append(type).Append('"');
            int start = marker.Ordinal ?? 1;
            if (ordered && start != 1) _builder.Append(" start=\"").Append(start.ToString(CultureInfo.InvariantCulture)).Append('"');
            _builder.Append('>');
            _open.Add(new OpenList(level, marker.Format, start));
        }

        private void CloseLast() {
            OpenList list = _open[_open.Count - 1];
            _open.RemoveAt(_open.Count - 1);
            _builder.Append("</li>").Append(list.Format == PdfCore.PdfLogicalListNumberFormat.Bullet ? "</ul>" : "</ol>");
            if (_open.Count == 0) _builder.AppendLine();
        }

        private sealed class OpenList {
            internal OpenList(int level, PdfCore.PdfLogicalListNumberFormat format, int next) {
                Level = level;
                Format = format;
                Next = next;
            }

            internal int Level { get; }

            internal PdfCore.PdfLogicalListNumberFormat Format { get; }

            internal int Next { get; set; }
        }
    }
}
