namespace OfficeIMO.IWork.Internal;

internal sealed partial class IWorkPagesArchiveBuilder {
    private readonly Dictionary<IWorkPagesNumberedList, ulong> _numberedListStyles = new();
    private readonly IWorkDecimalCreationCounter _encodedCounter = new();
    private IWorkPagesNumberedList? _previousNumberedList;

    private ulong AddNumberedListStyle(IWorkPagesNumberedList list) {
        if (_numberedListStyles.TryGetValue(list, out ulong id)) return id;
        id = Allocate(); _numberedListStyles.Add(list, id);
        var style = P().Message(1, Style(id, "OfficeIMO Decimal List")).UInt(10, 9)
            .Bool(18, true).Bool(20, true).Bool(22, false).String(23, list.FontName);
        // Native list vectors carry all nine slots. Retain the established two-level
        // graph for existing models; enable deeper slots when they are authored.
        // At scale one the first retained character run determines marker size.
        int numberedLevels = Math.Max(2, list.Items.Max(item => item.ListLevel) + 1);
        for (int level = 0; level < 9; level++) {
            style.UInt(11, level < numberedLevels ? 3u : 0u).Float(12, level < numberedLevels ? 1.5f : 0)
                .Float(13, level < numberedLevels ? level * 18 : 0)
                .Message(14, P().Float(1, 1).Float(2, 0).Bool(3, true))
                .UInt(15, (uint)list.MarkerStyle).Bool(25, false);
        }
        Add(id, 2023, style); return id;
    }

    private int EncodeListStart(IWorkPagesParagraphData paragraph) {
        bool changed = _previousNumberedList != paragraph.NumberedList;
        if (changed) { _encodedCounter.Reset(); _previousNumberedList = paragraph.NumberedList; }
        int number = paragraph.ListNumber!.Value;
        // Ordinary continuation remains native continuation. After switching lists,
        // restore each level only when its owned counter differs from the native scope.
        int start = changed || paragraph.ListRestartNumber.HasValue || _encodedCounter.Next(paragraph.ListLevel) != number
            ? number : 0;
        _encodedCounter.Accept(paragraph.ListLevel, number);
        return start;
    }
}
