using OfficeIMO.IWork.Internal;

namespace OfficeIMO.IWork;

/// <summary>Owns a decimal list whose items occupy separate paragraphs in one newly created Pages document.</summary>
/// <remarks>Items use two non-tiered levels, black text and a shared whole-point font size.
/// Plain paragraphs and attachments preserve counters. Returning to this list after another list resumes its counters.
/// Advancing or restarting a parent resets its child counter. Instances are not thread safe.</remarks>
public sealed class IWorkPagesNumberedList {
    private readonly IWorkPagesDocument _document;
    private readonly IWorkDecimalCreationCounter _counter = new();
    private readonly List<IWorkPagesParagraphData> _items = new();

    internal IWorkPagesNumberedList(IWorkPagesDocument document, string fontName, float fontSizePoints) {
        // Reuse character validation before exposing a list or mutating its document.
        var font = new IWorkTextRunData(string.Empty, fontName, fontSizePoints);
        if (fontSizePoints != Math.Round(fontSizePoints))
            throw new ArgumentOutOfRangeException(nameof(fontSizePoints), "Decimal lists require a whole-point font size.");
        _document = document; FontName = font.FontName; FontSizePoints = font.FontSizePoints;
        Items = _items.AsReadOnly();
    }

    /// <summary>Gets the shared body and marker font.</summary>
    public string FontName { get; }
    /// <summary>Gets the shared whole-point body and marker size.</summary>
    public float FontSizePoints { get; }
    /// <summary>Gets this list's items in their body insertion order.</summary>
    public IReadOnlyList<IWorkPagesParagraphData> Items { get; }

    /// <summary>Adds one numbered body paragraph, continuing its level unless a positive restart is supplied.</summary>
    /// <param name="text">Item text without line breaks. TAB text requires an explicit layout.</param>
    /// <param name="level">Zero for a parent or one for a child. A child requires an earlier parent in this list.</param>
    /// <param name="restartAt">Optional positive counter value. The first parent starts at one by default.</param>
    /// <param name="layout">Optional immutable paragraph spacing and tab stops.</param>
    public IWorkPagesParagraphData AddItem(string text, int level = 0, int? restartAt = null,
        IWorkPagesParagraphLayout? layout = null) {
        if (level is < 0 or > 1) throw new ArgumentOutOfRangeException(nameof(level));
        if (restartAt <= 0) throw new ArgumentOutOfRangeException(nameof(restartAt));
        if (level == 1 && !_counter.HasParent)
            throw new InvalidOperationException("A child item requires an earlier parent in this list.");
        int number = restartAt ?? _counter.Next(level)
            ?? throw new InvalidOperationException("The decimal counter requires an explicit restart after Int32.MaxValue.");
        var run = new IWorkTextRunData(text, FontName, FontSizePoints,
            IWorkCreationValidation.ParseColor("000000"), allowTabs: layout != null);
        var paragraph = new IWorkPagesParagraphData(new[] { run }, layout, this, level, number, restartAt);
        _document.AddNumberedParagraph(paragraph); _items.Add(paragraph); _counter.Accept(level, number);
        return paragraph;
    }
}
