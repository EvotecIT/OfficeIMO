using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    public static partial class WordDocumentTraversal {
        private const int MaxResolvedListTabStops = 1024;

        private static Tabs? CloneListLevelTabStops(Level level) {
            Tabs? tabs = level.GetFirstChild<PreviousParagraphProperties>()?.GetFirstChild<Tabs>();
            return tabs == null ? null : new Tabs(tabs.Elements<TabStop>().Take(MaxResolvedListTabStops)
                .Select(tab => (TabStop)tab.CloneNode(true)));
        }
    }
}
