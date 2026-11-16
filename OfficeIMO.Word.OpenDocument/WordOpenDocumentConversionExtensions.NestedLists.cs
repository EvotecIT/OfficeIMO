using OfficeIMO.OpenDocument;

namespace OfficeIMO.Word.OpenDocument;

public static partial class WordOpenDocumentConversionExtensions {
    /// <summary>
    /// Writes consecutive Word list paragraphs as ODT lists, nesting a deeper item inside the item before it. A
    /// change between bullets and numbers at a level, or an ordered counter that starts again, begins a new list.
    /// </summary>
    private sealed class OdtNestedLists {
        private const int MaximumDepth = 10;
        private readonly OdtDocument _target;
        private readonly List<OpenList> _open = new();

        internal OdtNestedLists(OdtDocument target) {
            _target = target;
        }

        internal void Close() => _open.Clear();

        internal OdtListItem AddItem(WordParagraphInfo paragraph, ref int lists, ref int flattenedLevels) {
            bool ordered = paragraph.IsOrderedList == true;
            int sourceLevel = Math.Max(0, paragraph.ListLevel ?? 0);
            int level = Math.Min(MaximumDepth - 1, sourceLevel);
            if (level != sourceLevel) flattenedLevels++;

            while (_open.Count > 0 && _open[_open.Count - 1].Level > level) _open.RemoveAt(_open.Count - 1);
            if (_open.Count > 0 && _open[_open.Count - 1].Level == level) {
                OpenList current = _open[_open.Count - 1];
                bool restarted = ordered && paragraph.ListIndex.HasValue && current.LastIndex.HasValue &&
                    paragraph.ListIndex.Value != current.LastIndex.Value + 1;
                if (current.Ordered != ordered || restarted) _open.RemoveAt(_open.Count - 1);
            }

            if (_open.Count == 0 || _open[_open.Count - 1].Level < level) {
                OdtList list = _open.Count == 0
                    ? _target.AddList(ordered)
                    : _open[_open.Count - 1].LastItem!.AddList(ordered);
                if (_open.Count == 0) lists++;
                _open.Add(new OpenList(level, ordered, list));
            }

            OpenList owner = _open[_open.Count - 1];
            OdtListItem item = owner.List.AddItem();
            owner.LastItem = item;
            owner.LastIndex = paragraph.ListIndex;
            return item;
        }

        private sealed class OpenList {
            internal OpenList(int level, bool ordered, OdtList list) {
                Level = level;
                Ordered = ordered;
                List = list;
            }

            internal int Level { get; }

            internal bool Ordered { get; }

            internal OdtList List { get; }

            internal OdtListItem? LastItem { get; set; }

            internal int? LastIndex { get; set; }
        }
    }
}
