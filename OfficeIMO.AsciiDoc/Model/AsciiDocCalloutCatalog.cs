namespace OfficeIMO.AsciiDoc;

/// <summary>Correlation between verbatim callout markers and their following explanation lists.</summary>
public sealed class AsciiDocCalloutCatalog {
    private readonly List<AsciiDocCalloutGroup> _groups = new List<AsciiDocCalloutGroup>();
    private readonly List<AsciiDocReferenceDiagnostic> _diagnostics = new List<AsciiDocReferenceDiagnostic>();
    private readonly IReadOnlyList<AsciiDocCalloutGroup> _groupView;
    private readonly IReadOnlyList<AsciiDocReferenceDiagnostic> _diagnosticView;
    private AsciiDocCalloutCatalog() { _groupView = _groups.AsReadOnly(); _diagnosticView = _diagnostics.AsReadOnly(); }
    /// <summary>Callout groups in source order.</summary>
    public IReadOnlyList<AsciiDocCalloutGroup> Groups => _groupView;
    /// <summary>Missing, duplicate, or unmatched explanations.</summary>
    public IReadOnlyList<AsciiDocReferenceDiagnostic> Diagnostics => _diagnosticView;
    /// <summary>Builds a bounded snapshot including compound blocks and AsciiDoc-style table cells.</summary>
    public static AsciiDocCalloutCatalog Create(AsciiDocDocument document, int maximumNestingDepth = 64, System.Threading.CancellationToken cancellationToken = default) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        if (maximumNestingDepth < 1) throw new ArgumentOutOfRangeException(nameof(maximumNestingDepth));
        var result = new AsciiDocCalloutCatalog(); result.Visit(document, 0, maximumNestingDepth, cancellationToken); return result;
    }
    private void Visit(AsciiDocDocument document, int depth, int maximumDepth, System.Threading.CancellationToken token) {
        if (depth >= maximumDepth) throw new InvalidDataException("AsciiDoc callout traversal exceeds MaximumNestingDepth.");
        AsciiDocBlock? previous = null;
        foreach (AsciiDocBlock block in document.Blocks) {
            token.ThrowIfCancellationRequested();
            if (block is AsciiDocBlankLine || block is AsciiDocLineComment || block is IAsciiDocBlockMetadata) continue;
            if (block is AsciiDocListBlock list && list.Kind == AsciiDocListKind.Callout) {
                AsciiDocDelimitedBlock? code = previous as AsciiDocDelimitedBlock;
                if (code == null || code.Kind != AsciiDocDelimitedBlockKind.Listing && code.Kind != AsciiDocDelimitedBlockKind.Literal) {
                    Report("ADOCCO001", "Callout explanations have no preceding listing or literal block.", list.Span);
                } else AddGroup(code, list, token);
            }
            if (block is AsciiDocDelimitedBlock compound && compound.GetBody(token) is AsciiDocDocument body) Visit(body, depth + 1, maximumDepth, token);
            if (block is AsciiDocTableBlock table) foreach (AsciiDocTableCell cell in table.Table.Cells)
                if (cell.GetBody(token) is AsciiDocDocument cellBody) Visit(cellBody, depth + 1, maximumDepth, token);
            previous = block;
        }
    }
    private void AddGroup(AsciiDocDelimitedBlock code, AsciiDocListBlock list, System.Threading.CancellationToken token) {
        var markers = new Dictionary<int, List<int>>();
        int automatic = 0;
        string content = code.Content;
        for (int offset = 0; offset < content.Length; offset++) {
            if ((offset & 1023) == 0) token.ThrowIfCancellationRequested();
            if (content[offset] == '\\') { offset++; continue; }
            if (content[offset] != '<') continue;
            int end = offset + 1;
            while (end < content.Length && end - offset <= 10 && content[end] >= '0' && content[end] <= '9') end++;
            if (end == offset + 1 && end < content.Length && content[end] == '.') end++;
            if (end >= content.Length || content[end] != '>' || end == offset + 1) continue;
            string value = content.Substring(offset + 1, end - offset - 1);
            int number;
            if (value == ".") number = ++automatic;
            else if (!int.TryParse(value, out number) || number < 1) continue;
            if (!markers.TryGetValue(number, out List<int>? positions)) { positions = new List<int>(); markers.Add(number, positions); }
            positions.Add(offset);
            offset = end;
        }
        automatic = 0;
        var entries = new List<AsciiDocCallout>();
        var explained = new HashSet<int>();
        foreach (AsciiDocListItem item in list.Items) {
            token.ThrowIfCancellationRequested();
            int number = item.Marker == "<.>" ? ++automatic : int.Parse(item.Marker.Substring(1, item.Marker.Length - 2), System.Globalization.CultureInfo.InvariantCulture);
            if (!explained.Add(number)) Report("ADOCCO002", "Duplicate callout explanation " + number + ".", item.Syntax.Span);
            if (!markers.TryGetValue(number, out List<int>? positions)) Report("ADOCCO003", "Callout explanation " + number + " has no code marker.", item.Syntax.Span);
            entries.Add(new AsciiDocCallout(number, item, (positions ?? new List<int>()).AsReadOnly()));
        }
        foreach (int number in markers.Keys.Where(number => !explained.Contains(number))) Report("ADOCCO004", "Code callout " + number + " has no explanation.", code.Span);
        _groups.Add(new AsciiDocCalloutGroup(code, list, entries.AsReadOnly()));
    }
    private void Report(string code, string message, AsciiDocSourceSpan span) => _diagnostics.Add(new AsciiDocReferenceDiagnostic(code, message, span));
}

/// <summary>A verbatim block and its following callout explanation list.</summary>
public sealed class AsciiDocCalloutGroup {
    internal AsciiDocCalloutGroup(AsciiDocDelimitedBlock code, AsciiDocListBlock list, IReadOnlyList<AsciiDocCallout> items) { CodeBlock = code; ExplanationList = list; Items = items; }
    /// <summary>Source listing or literal block.</summary>
    public AsciiDocDelimitedBlock CodeBlock { get; }
    /// <summary>Source explanation list.</summary>
    public AsciiDocListBlock ExplanationList { get; }
    /// <summary>Numbered explanations and their marker positions.</summary>
    public IReadOnlyList<AsciiDocCallout> Items { get; }
}
/// <summary>A numbered explanation correlated with zero or more verbatim markers.</summary>
public sealed class AsciiDocCallout {
    internal AsciiDocCallout(int number, AsciiDocListItem explanation, IReadOnlyList<int> offsets) { Number = number; Explanation = explanation; MarkerOffsets = offsets; }
    /// <summary>Resolved callout number.</summary>
    public int Number { get; }
    /// <summary>Typed explanation text.</summary>
    public AsciiDocListItem Explanation { get; }
    /// <summary>Zero-based offsets into the code block's current Content, rather than the original file.</summary>
    public IReadOnlyList<int> MarkerOffsets { get; }
}
