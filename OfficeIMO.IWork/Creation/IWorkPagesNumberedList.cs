using OfficeIMO.IWork.Internal;

namespace OfficeIMO.IWork;

/// <summary>Owns a decimal list whose items occupy separate paragraphs in one newly created Pages document.</summary>
/// <remarks>Items use nine non-tiered levels and whole-point font sizes. Markers remain plain black in the list's font family;
/// their size follows the first character run. Rich runs can vary their size, font family, opaque color and character decorations.
/// Plain paragraphs and attachments preserve counters. Returning to this list after another list resumes its counters.
/// Advancing or restarting a parent resets all its descendant counters. Instances are not thread safe.</remarks>
public sealed class IWorkPagesNumberedList {
    private readonly IWorkPagesDocument _document;
    private readonly IWorkDecimalCreationCounter _counter = new();
    private readonly List<IWorkPagesParagraphData> _items = new();

    internal IWorkPagesNumberedList(IWorkPagesDocument document, string fontName, float fontSizePoints,
        IWorkDecimalListMarkerStyle markerStyle) {
        // Reuse character validation before exposing a list or mutating its document.
        var font = new IWorkTextRunData(string.Empty, fontName, fontSizePoints);
        if (fontSizePoints != Math.Round(fontSizePoints))
            throw new ArgumentOutOfRangeException(nameof(fontSizePoints), "Decimal lists require a whole-point font size.");
        if ((uint)markerStyle > (uint)IWorkDecimalListMarkerStyle.ClosingParenthesis)
            throw new ArgumentOutOfRangeException(nameof(markerStyle));
        _document = document; FontName = font.FontName; FontSizePoints = font.FontSizePoints;
        MarkerStyle = markerStyle;
        Items = _items.AsReadOnly();
    }

    /// <summary>Gets the marker font and default body font.</summary>
    public string FontName { get; }
    /// <summary>Gets the default whole-point body size, also used for paragraph-relative indentation.</summary>
    public float FontSizePoints { get; }
    /// <summary>Gets the decimal marker punctuation shared by every level of this list.</summary>
    public IWorkDecimalListMarkerStyle MarkerStyle { get; }
    /// <summary>Gets this list's items in their body insertion order.</summary>
    public IReadOnlyList<IWorkPagesParagraphData> Items { get; }

    /// <summary>Adds one numbered body paragraph, continuing its level unless a positive restart is supplied.</summary>
    /// <param name="text">Item text without line breaks. TAB text requires an explicit layout.</param>
    /// <param name="level">Zero-based native level from zero through eight. A nested item requires an earlier parent at the preceding level in this list.</param>
    /// <param name="restartAt">Optional positive counter value. The first parent starts at one by default.</param>
    /// <param name="layout">Optional immutable paragraph spacing and tab stops.</param>
    public IWorkPagesParagraphData AddItem(string text, int level = 0, int? restartAt = null,
        IWorkPagesParagraphLayout? layout = null) {
        var run = new IWorkTextRunData(text, FontName, FontSizePoints,
            IWorkCreationValidation.ParseColor("000000"), allowTabs: layout != null);
        return AppendItem(new[] { run }, level, restartAt, layout);
    }

    /// <summary>Adds a snapshot of rich character runs as one numbered body paragraph.</summary>
    /// <param name="runs">Runs with whole-point sizes. Sizes, font families, opaque colors and character decorations may vary.</param>
    /// <param name="level">Zero-based native level from zero through eight. A nested item requires an earlier parent at the preceding level in this list.</param>
    /// <param name="restartAt">Optional positive counter value. The first parent starts at one by default.</param>
    /// <param name="layout">Optional immutable paragraph spacing and tab stops.</param>
    /// <remarks>Runs cannot contain line or paragraph breaks. Empty ranges are discarded unless the whole item is blank.
    /// The marker uses the first retained run's size and the list's font family and punctuation.</remarks>
    public IWorkPagesParagraphData AddRichItem(IEnumerable<IWorkTextRunData> runs, int level = 0, int? restartAt = null,
        IWorkPagesParagraphLayout? layout = null) => AppendItem(runs, level, restartAt, layout);

    private IWorkPagesParagraphData AppendItem(IEnumerable<IWorkTextRunData> runs, int level, int? restartAt,
        IWorkPagesParagraphLayout? layout) {
        if (level is < 0 or > IWorkDecimalListProfile.MaximumLevel) throw new ArgumentOutOfRangeException(nameof(level));
        if (restartAt <= 0) throw new ArgumentOutOfRangeException(nameof(restartAt));
        if (level > 0 && !_counter.HasParent(level))
            throw new InvalidOperationException("A nested item requires an earlier parent at the preceding level in this list.");
        int number = restartAt ?? _counter.Next(level)
            ?? throw new InvalidOperationException("The decimal counter requires an explicit restart after Int32.MaxValue.");
        var paragraph = new IWorkPagesParagraphData(runs, layout, this, level, number, restartAt);
        if (paragraph.Runs.Any(run => run.FontSizePoints != Math.Round(run.FontSizePoints)))
            throw new ArgumentException("Rich numbered item runs require whole-point font sizes.", nameof(runs));
        _document.AddNumberedParagraph(paragraph); _items.Add(paragraph); _counter.Accept(level, number);
        return paragraph;
    }
}
