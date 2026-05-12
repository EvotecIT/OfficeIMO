using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Text;

namespace OfficeIMO.Excel {
    internal sealed class SharedStringCache {
        private readonly List<string> _items;

        private SharedStringCache(List<string> items) => _items = items;

        public static SharedStringCache Build(SpreadsheetDocument doc) {
            var part = doc.WorkbookPart!.SharedStringTablePart;
            if (part?.SharedStringTable == null) return new SharedStringCache(new List<string>());

            var list = new List<string>(Math.Max(1024, (int)(part.SharedStringTable.Count?.Value ?? 0)));
            foreach (var item in part.SharedStringTable.Elements<SharedStringItem>()) {
                if (item.Text?.Text != null)
                    list.Add(item.Text.Text);
                else if (item.HasChildren)
                    list.Add(GetRunText(item));
                else
                    list.Add(string.Empty);
            }
            return new SharedStringCache(list);
        }

        internal static string GetRunText(OpenXmlElement parent) {
            string? first = null;
            StringBuilder? builder = null;

            foreach (var run in parent.Elements<Run>()) {
                string text = run.Text?.Text ?? string.Empty;
                if (builder != null) {
                    builder.Append(text);
                } else if (first == null) {
                    first = text;
                } else {
                    builder = new StringBuilder(first.Length + text.Length);
                    builder.Append(first);
                    builder.Append(text);
                }
            }

            return builder?.ToString() ?? first ?? string.Empty;
        }

        public string? Get(int index) {
            if ((uint)index < (uint)_items.Count) return _items[index];
            return null;
        }
    }
}

