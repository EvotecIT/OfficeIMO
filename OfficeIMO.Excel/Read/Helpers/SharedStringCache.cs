using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Text;
using System.Threading;

namespace OfficeIMO.Excel {
    internal sealed class SharedStringCache {
        private const int MinimumInitialCapacity = 1024;

        private readonly SharedStringTablePart? _part;
        private readonly Lazy<List<string>> _items;

        private SharedStringCache(SharedStringTablePart? part) {
            _part = part;
            _items = new Lazy<List<string>>(LoadItems, LazyThreadSafetyMode.ExecutionAndPublication);
        }

        public static SharedStringCache Build(SpreadsheetDocument doc) {
            return new SharedStringCache(doc.WorkbookPart!.SharedStringTablePart);
        }

        private List<string> LoadItems() {
            if (_part?.SharedStringTable == null) return new List<string>();

            var table = _part.SharedStringTable;
            var list = new List<string>(Math.Max(MinimumInitialCapacity, (int)(table.Count?.Value ?? 0)));
            foreach (var item in table.Elements<SharedStringItem>()) {
                if (item.Text?.Text != null)
                    list.Add(item.Text.Text);
                else if (item.HasChildren)
                    list.Add(GetRunText(item));
                else
                    list.Add(string.Empty);
            }

            return list;
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
            var items = _items.Value;
            if ((uint)index < (uint)items.Count) return items[index];
            return null;
        }
    }
}

