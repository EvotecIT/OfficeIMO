using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel
{
    internal sealed class SharedStringCache
    {
        private readonly List<string> _items;

        private SharedStringCache(List<string> items) => _items = items;

        public static SharedStringCache Build(SpreadsheetDocument doc)
        {
            var part = doc.WorkbookPart!.SharedStringTablePart;
            if (part?.SharedStringTable == null) return new SharedStringCache(new List<string>());

            var list = new List<string>(Math.Max(1024, (int)(part.SharedStringTable.Count?.Value ?? 0)));
            foreach (var item in part.SharedStringTable.Elements<SharedStringItem>())
            {
                if (item.Text?.Text != null)
                    list.Add(item.Text.Text);
                else if (item.HasChildren)
                    list.Add(string.Concat(item.Elements<Run>().Select(r => r.Text?.Text ?? string.Empty)));
                else
                    list.Add(string.Empty);
            }
            return new SharedStringCache(list);
        }

        public string? Get(int index)
        {
            if ((uint)index < (uint)_items.Count) return _items[index];
            return null;
        }
    }
}

