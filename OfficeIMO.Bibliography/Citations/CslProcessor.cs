using System.Text.Json;
using System.Xml.Linq;

namespace OfficeIMO.Bibliography;

/// <summary>A managed CSL processor over an immutable citation-data snapshot. Rendering has no filesystem, network, or executable dependency.</summary>
public sealed partial class CslProcessor {
    private readonly CslStyle _style;
    private readonly CslRenderOptions _options;
    private readonly Dictionary<string, CslRecord> _records;

    /// <summary>Fidelity evidence for projecting the source bibliography into the processor's CSL data snapshot.</summary>
    /// <remarks>Rendering preserves this evidence independently of style formatting. Set <see cref="CslRenderOptions.RequireNoDataLoss"/> to reject an approximate data projection.</remarks>
    public BibliographyConversionReport DataConversionReport { get; }

    /// <summary>Creates a processor by copying the bibliography data and rendering options. Keys must be nonempty and unique.</summary>
    public CslProcessor(BibliographyDocument document, CslStyle style, CslRenderOptions? options = null, CancellationToken cancellationToken = default) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        _style = style ?? throw new ArgumentNullException(nameof(style));
        _options = CopyOptions(options ?? new CslRenderOptions());
        if (document.Items.Any(item => string.IsNullOrWhiteSpace(item.Key)) || document.Items.Select(item => item.Key).Distinct(StringComparer.Ordinal).Count() != document.Items.Count)
            throw new ArgumentException("CSL rendering requires nonempty, unique bibliography keys.", nameof(document));
        BibliographyWriteResult written = document.Write(new BibliographyWriteOptions { Format = BibliographyFormat.CslJson, Mode = BibliographyWriterMode.Canonical, RequireNoLoss = _options.RequireNoDataLoss }, cancellationToken);
        DataConversionReport = written.Report;
        using JsonDocument json = JsonDocument.Parse(written.Content);
        JsonElement[] data = json.RootElement.ValueKind == JsonValueKind.Array ? json.RootElement.EnumerateArray().ToArray() : new[] { json.RootElement };
        _records = new Dictionary<string, CslRecord>(StringComparer.Ordinal);
        for (int index = 0; index < data.Length; index++) {
            cancellationToken.ThrowIfCancellationRequested();
            var record = new CslRecord(data[index], index + 1);
            _records.Add(record.Key, record);
        }
    }

    /// <summary>Renders a complete document citation sequence, recalculating numbers, positions, disambiguation, and bibliography order.</summary>
    /// <remarks>The processor does not mutate the source bibliography. Replacing the input sequence handles insertion, deletion, and movement of citations.</remarks>
    public CslRenderResult Render(IEnumerable<CslCitation> citations, bool includeUncitedItems = false, CancellationToken cancellationToken = default) {
        if (citations == null) throw new ArgumentNullException(nameof(citations));
        // Operation-local copies keep successive operations independent and allow
        // callers to render different documents concurrently with one processor.
        Dictionary<string, CslRecord> records = _records.ToDictionary(pair => pair.Key, pair => new CslRecord(pair.Value), StringComparer.Ordinal);
        CslCitation[] clusters = CopyCitations(citations, records, cancellationToken);
        var locale = new CslLocale(_style, _options, cancellationToken);
        var evaluator = new CslEvaluator(_style, locale, _options, cancellationToken);
        CslRecord[] ordered = AssignNumbers(records, clusters, includeUncitedItems);
        if (!_style.IsNoteStyle) foreach (CslRecord record in ordered) record.FirstNote = 0;
        XElement? bibliography = _style.Root.Element(CslStyle.Namespace + "bibliography");
        if (bibliography != null) ordered = Sort(ordered, bibliography, evaluator, XElementScope.Bibliography);
        for (int index = 0; index < ordered.Length; index++) ordered[index].Number = index + 1;
        Disambiguate(ordered, clusters, evaluator);
        var citationOutput = new List<CslRenderedEntry>();
        var bibliographyOutput = new List<CslRenderedEntry>();
        int outputLength = 0;
        int maximumLeftMarginCharacters = 0;
        XElement citation = _style.Root.Element(CslStyle.Namespace + "citation")!;
        int near = int.TryParse((string?)citation.Attribute("near-note-distance"), out int configured) ? configured : 5;
        var positions = new CslPositionHistory(_style.IsNoteStyle, near);
        XElement layout = citation.Element(CslStyle.Namespace + "layout")!;
        for (int clusterIndex = 0; clusterIndex < clusters.Length; clusterIndex++) {
            cancellationToken.ThrowIfCancellationRequested();
            CslCitation cluster = clusters[clusterIndex];
            CslCitationItem[] items = SortCites(cluster.Items, records, citation, evaluator);
            var values = new List<CiteRendering>();
            CslCitationItem? preceding = positions.Preceding(cluster);
            foreach (CslCitationItem item in items) {
                CslRecord record = records[item.Key];
                var context = CreateContext(record, XElementScope.Citation);
                context.Cite = item;
                positions.SetPosition(context, cluster, preceding);
                CslText value;
                if (item.AuthorOnly) {
                    evaluator.Evaluate(ItemLayout(layout), context);
                    value = context.NarrativeNames ?? evaluator.Evaluate(new XElement(CslStyle.Namespace + "names",
                        new XAttribute("variable", "author"), new XElement(CslStyle.Namespace + "name")), context);
                } else {
                    XElement itemLayout = ItemLayout(layout);
                    value = evaluator.Evaluate(itemLayout, context);
                }
                values.Add(new CiteRendering(item, context, value.Affix(item.Prefix ?? string.Empty, item.Suffix ?? string.Empty)));
                preceding = item;
            }
            CslText rendered = Collapse(values, citation, layout, evaluator);
            if (!items.All(item => item.AuthorOnly)) rendered = rendered.Decorate(layout, locale, cancellationToken: cancellationToken);
            CslCitationItem? firstVisible = values.FirstOrDefault(value => !value.Value.IsEmpty)?.Item;
            if (positions.StartsNote(cluster, firstVisible, rendered)) rendered = rendered.CapitalizeNoteStart(locale.Culture, cancellationToken);
            positions.Complete(cluster);
            string content = Output(rendered, locale, cancellationToken, out _, out bool isEmpty);
            CheckOutput(ref outputLength, content.Length);
            citationOutput.Add(new CslRenderedEntry(cluster.Id, content, isEmpty));
        }
        if (bibliography?.Element(CslStyle.Namespace + "layout") is XElement bibliographyLayout) {
            string[] previousNames = Array.Empty<string>();
            foreach (CslRecord record in ordered) {
                cancellationToken.ThrowIfCancellationRequested();
                CslContext context = CreateContext(record, XElementScope.Bibliography);
                context.PreviousNames = previousNames;
                context.NamesReplacement = (string?)bibliography.Attribute("subsequent-author-substitute");
                context.NamesReplacementRule = (string?)bibliography.Attribute("subsequent-author-substitute-rule") ?? "complete-all";
                CslText rendered = evaluator.Evaluate(bibliographyLayout, context);
                previousNames = context.FirstNames;
                string content = Output(rendered, locale, cancellationToken, out int leftMarginCharacters, out bool isEmpty);
                maximumLeftMarginCharacters = Math.Max(maximumLeftMarginCharacters, leftMarginCharacters);
                if (_options.OutputFormat == CslOutputFormat.Html) content = "<div class=\"csl-entry\">" + content + "</div>";
                CheckOutput(ref outputLength, content.Length);
                bibliographyOutput.Add(new CslRenderedEntry(record.Key, content, isEmpty));
            }
        }
        return new CslRenderResult(citationOutput.AsReadOnly(), bibliographyOutput.AsReadOnly(),
            bibliography == null ? null : new CslBibliographyLayout(bibliography, maximumLeftMarginCharacters));
    }

    /// <summary>Renders all records as a bibliography without requiring document citations.</summary>
    public IReadOnlyList<CslRenderedEntry> RenderBibliography(CancellationToken cancellationToken = default) => Render(Array.Empty<CslCitation>(), true, cancellationToken).Bibliography;
    private string Output(CslText value, CslLocale locale, CancellationToken token, out int maximumLeftMarginCharacters, out bool isEmpty) {
        value = value.ResolveQuotes(locale, _options.MaximumIntermediateCharacters, token)
            .ProjectDisplay(_options.MaximumIntermediateCharacters, token)
            .LocalizePunctuation(locale, _options.MaximumIntermediateCharacters, token);
        isEmpty = value.IsEmpty;
        if (_options.OutputFormat == CslOutputFormat.Html)
            return value.FinalizeHtml(_options.MaximumIntermediateCharacters, token, out maximumLeftMarginCharacters);
        maximumLeftMarginCharacters = value.MeasureLeftMarginCharacters(_options.MaximumIntermediateCharacters, token);
        return value.Plain;
    }
    private CslContext CreateContext(CslRecord record, XElementScope scope) => new CslContext(record, scope) { AutomaticYearSuffix = !_style.HasExplicitYearSuffix };
    private CslText Join(IEnumerable<CslText> values, string delimiter) => CslText.Join(values, delimiter, maximumCharacters: _options.MaximumIntermediateCharacters);
    private static XElement ItemLayout(XElement layout) => new XElement(CslStyle.Namespace + "layout", layout.Elements().Select(CslElementIdentity.Copy));
    private void CheckOutput(ref int current, int length) {
        if ((long)current + length > _options.MaximumOutputCharacters) throw new InvalidDataException("CSL rendering exceeds MaximumOutputCharacters.");
        current += length;
    }

    private static CslRenderOptions CopyOptions(CslRenderOptions options) {
        if (options.MaximumOutputCharacters < 1 || options.MaximumIntermediateCharacters < 1 || options.MaximumCitationItems < 1 || options.MaximumRenderingOperations < 1 || !Enum.IsDefined(typeof(CslOutputFormat), options.OutputFormat)) throw new ArgumentOutOfRangeException(nameof(options));
        var copy = new CslRenderOptions { OutputFormat = options.OutputFormat, LinkBibliographyIdentifiers = options.LinkBibliographyIdentifiers, Locale = options.Locale, MaximumOutputCharacters = options.MaximumOutputCharacters, MaximumIntermediateCharacters = options.MaximumIntermediateCharacters, MaximumCitationItems = options.MaximumCitationItems, MaximumRenderingOperations = options.MaximumRenderingOperations, RequireNoDataLoss = options.RequireNoDataLoss };
        foreach (var locale in options.Locales) copy.Locales.Add(locale.Key, locale.Value);
        foreach (var variable in options.Abbreviations) copy.Abbreviations.Add(variable.Key, new Dictionary<string, string>(variable.Value, StringComparer.Ordinal));
        return copy;
    }

    private CslCitation[] CopyCitations(IEnumerable<CslCitation> citations, IDictionary<string, CslRecord> records, CancellationToken token) {
        var result = new List<CslCitation>();
        var ids = new HashSet<string>(StringComparer.Ordinal);
        int itemCount = 0;
        foreach (CslCitation source in citations) {
            token.ThrowIfCancellationRequested();
            if (source == null || !ids.Add(source.Id) || source.NoteIndex < 0) throw new ArgumentException("Citation clusters require unique identifiers and nonnegative note numbers.", nameof(citations));
            var copy = new CslCitation(source.Id) { NoteIndex = source.NoteIndex, NoteHasPrecedingText = source.NoteHasPrecedingText };
            foreach (CslCitationItem item in source.Items) {
                if (item == null || !records.ContainsKey(item.Key)) throw new ArgumentException("Citation refers to an unknown bibliography key.", nameof(citations));
                if (item.AuthorOnly && item.SuppressAuthor) throw new ArgumentException("AuthorOnly and SuppressAuthor cannot both be enabled for one citation item.", nameof(citations));
                if (!string.IsNullOrEmpty(item.Locator) && string.IsNullOrWhiteSpace(item.LocatorType))
                    throw new ArgumentException("A nonempty locator requires a locator type.", nameof(citations));
                if (++itemCount > _options.MaximumCitationItems) throw new InvalidDataException("CSL citations exceed MaximumCitationItems.");
                copy.Items.Add(new CslCitationItem(item.Key) { Locator = item.Locator, LocatorType = item.LocatorType, Prefix = item.Prefix, Suffix = item.Suffix, AuthorOnly = item.AuthorOnly, SuppressAuthor = item.SuppressAuthor });
            }
            result.Add(copy);
            if (result.Count > _options.MaximumCitationItems) throw new InvalidDataException("CSL clusters exceed MaximumCitationItems.");
        }
        return result.ToArray();
    }

    private static CslRecord[] AssignNumbers(IDictionary<string, CslRecord> records, IEnumerable<CslCitation> citations, bool uncited) {
        var ordered = new List<CslRecord>();
        var seen = new HashSet<string>(StringComparer.Ordinal);
        foreach (CslCitation citation in citations) foreach (CslCitationItem cite in citation.Items) {
            CslRecord record = records[cite.Key];
            if (seen.Add(cite.Key)) ordered.Add(record);
            if (record.FirstNote == 0 && citation.NoteIndex > 0) record.FirstNote = citation.NoteIndex;
        }
        if (uncited) ordered.AddRange(records.Values.Where(record => seen.Add(record.Key)).OrderBy(record => record.Number));
        for (int index = 0; index < ordered.Count; index++) ordered[index].Number = index + 1;
        return ordered.ToArray();
    }
}
