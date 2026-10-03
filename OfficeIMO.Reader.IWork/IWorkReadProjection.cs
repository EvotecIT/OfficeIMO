using OfficeIMO.IWork;
using OfficeIMO.IWork.Internal;

namespace OfficeIMO.Reader.IWork;

internal sealed partial class IWorkReadProjection {
    private const int MaximumMarkdownListLevel = 128;
    private readonly OfficeDocumentReadResult _result;
    private readonly string _path;
    private readonly ReaderOptions _readerOptions;
    private readonly ReaderIWorkOptions _options;
    private readonly CancellationToken _cancellationToken;
    private readonly IWorkProjectionBudget _projectionBudget;
    private readonly List<ReaderChunk> _chunks = new();
    private readonly List<OfficeDocumentBlock> _blocks = new();
    private readonly List<ReaderTable> _tables = new();
    private readonly List<OfficeDocumentAsset> _assets = new();
    private readonly List<OfficeDocumentLink> _links = new();
    private readonly List<OfficeDocumentPage> _pages = new();
    private readonly List<OfficeDocumentDiagnostic> _diagnostics = new();
    private readonly List<OfficeDocumentMetadataEntry> _projectionMetadata = new();
    private bool _reportedMarkdownListDepthLimit;
    private bool _reportedUnsupportedTextStyles;
    private bool _reportedUnsupportedLayoutBreaks;
    private bool _reportedTableBudgetExhausted;
    private int _projectedTableCells;
    private readonly Dictionary<OfficeDocumentPage, List<OfficeDocumentBlock>> _pageBlocks = new();
    private readonly Dictionary<OfficeDocumentPage, List<ReaderTable>> _pageTables = new();
    private readonly Dictionary<OfficeDocumentPage, List<OfficeDocumentAsset>> _pageAssets = new();
    private readonly Dictionary<OfficeDocumentPage, List<OfficeDocumentLink>> _pageLinks = new();

    internal IWorkReadProjection(OfficeDocumentReadResult result, string path,
        ReaderOptions readerOptions, ReaderIWorkOptions options,
        CancellationToken cancellationToken) {
        _result = result;
        _path = path;
        _readerOptions = readerOptions;
        _options = options;
        _cancellationToken = cancellationToken;
        _projectionBudget = new IWorkProjectionBudget((options.ReadOptions ?? new IWorkReadOptions()).Clone());
    }

    internal void AddNumbers(IWorkNumbersProjection source) {
        for (int sheetIndex = 0; sheetIndex < source.Sheets.Count; sheetIndex++) {
            _cancellationToken.ThrowIfCancellationRequested();
            IWorkNumbersSheet sheet = source.Sheets[sheetIndex];
            var page = NewPage(sheetIndex + 1, sheet.Name, sheet.Name);
            foreach (IWorkNumbersDrawable drawable in sheet.Drawables) {
                _cancellationToken.ThrowIfCancellationRequested();
                switch (drawable.Kind) {
                    case IWorkNumbersDrawableKind.TextBox:
                        AddPlainText(page, drawable.TextBox!, "text-box");
                        break;
                    case IWorkNumbersDrawableKind.Table:
                        AddTable(page, drawable.Table!);
                        break;
                }
            }
        }
        AddDiagnostics(source.Diagnostics);
    }

    internal void AddKeynote(IWorkKeynoteProjection source) {
        foreach (IWorkKeynoteSlide slide in source.Slides) {
            _cancellationToken.ThrowIfCancellationRequested();
            var page = NewPage(slide.Index, slide.Name, null, slide.Index);
            _projectionMetadata.Add(new OfficeDocumentMetadataEntry {
                Id = "iwork-slide-" + slide.Index.ToString("D4", CultureInfo.InvariantCulture)
                    + "-skipped",
                Category = "presentation.slide",
                Name = "IsSkipped",
                Value = slide.IsSkipped ? "true" : "false",
                ValueType = "boolean",
                Location = Location(page)
            });
            if (slide.HasBackgroundFill) {
                _diagnostics.Add(new OfficeDocumentDiagnostic {
                    Severity = OfficeDocumentDiagnosticSeverity.Warning,
                    Category = OfficeDocumentDiagnosticCategory.Content,
                    Source = "OfficeIMO.Reader.IWork",
                    Code = "IWORK_READER_SLIDE_BACKGROUND_OMITTED",
                    Message = "The plain Reader projection does not retain the Keynote slide background fill.",
                    Location = Location(page)
                });
            }
            page.Width = source.SlideSize?.WidthPoints;
            page.Height = source.SlideSize?.HeightPoints;
            foreach (IWorkKeynoteDrawable drawable in slide.Drawables) {
                _cancellationToken.ThrowIfCancellationRequested();
                switch (drawable.Kind) {
                    case IWorkKeynoteDrawableKind.TextBox:
                        AddTextBox(page, drawable.TextBox!,
                            drawable.IsTitlePlaceholder ? "title" : "text-box");
                        break;
                    case IWorkKeynoteDrawableKind.Image:
                        AddImage(page, drawable.Image!);
                        break;
                    case IWorkKeynoteDrawableKind.Table:
                        AddTable(page, drawable.Table!);
                        break;
                }
            }
            AddRichContent(page, slide.PresenterNoteContent, "presenter-notes");
        }
        AddDiagnostics(source.Diagnostics);
    }

    internal void Complete(IWorkSourceDocument source) {
        AddDiagnostics(source.Diagnostics);
        string[] warnings = _diagnostics
            .Where(diagnostic => diagnostic.Severity is OfficeDocumentDiagnosticSeverity.Warning
                or OfficeDocumentDiagnosticSeverity.Error)
            .Select(diagnostic => diagnostic.Code + ": " + diagnostic.Message)
            .Distinct(StringComparer.Ordinal).ToArray();
        if (warnings.Length > 0) {
            if (_chunks.Count == 0) {
                _chunks.Add(new ReaderChunk {
                    Id = "iwork-diagnostic-000",
                    Kind = ReaderInputKind.IWork,
                    Location = new ReaderLocation { Path = _path },
                    Warnings = warnings
                });
            } else {
                ReaderChunk first = _chunks[0];
                first.Warnings = (first.Warnings ?? Array.Empty<string>())
                    .Concat(warnings).Distinct(StringComparer.Ordinal).ToArray();
            }
        }
        _result.Chunks = _chunks.ToArray();
        _result.Blocks = _blocks.ToArray();
        _result.Tables = _tables.ToArray();
        _result.Assets = _assets.ToArray();
        _result.Links = _links.ToArray();
        _result.Pages = _pages.ToArray();
        _result.Diagnostics = _diagnostics.ToArray();
        foreach (OfficeDocumentPage page in _pages) {
            page.Blocks = _pageBlocks[page].ToArray();
            page.Tables = _pageTables[page].ToArray();
            page.Assets = _pageAssets[page].ToArray();
            page.Links = _pageLinks[page].ToArray();
        }
        var documentMarkdown = new StringBuilder();
        foreach (ReaderChunk chunk in _chunks) {
            if (string.IsNullOrEmpty(chunk.Markdown)) continue;
            if (documentMarkdown.Length > 0 && !chunk.ContinuesPreviousChunk) {
                documentMarkdown.Append("\n\n");
            }
            documentMarkdown.Append(chunk.Markdown);
        }
        _result.Markdown = documentMarkdown.ToString();
        _result.Metadata = source.BuildVersions.Select((version, index) =>
            new OfficeDocumentMetadataEntry {
                Id = "iwork-build-" + index.ToString("D4", CultureInfo.InvariantCulture),
                Category = "producer",
                Name = "BuildVersion",
                Value = version
            }).Concat(_projectionMetadata).ToArray();
    }

    private OfficeDocumentPage NewPage(int? number, string name, string? sheet,
        int? slide = null) {
        var page = new OfficeDocumentPage {
            Number = number,
            Name = name,
            Location = new ReaderLocation { Path = _path, Sheet = sheet, Slide = slide }
        };
        _pages.Add(page);
        _pageBlocks.Add(page, new List<OfficeDocumentBlock>());
        _pageTables.Add(page, new List<ReaderTable>());
        _pageAssets.Add(page, new List<OfficeDocumentAsset>());
        _pageLinks.Add(page, new List<OfficeDocumentLink>());
        return page;
    }

    private void AddRichContent(OfficeDocumentPage page, IWorkTextContent content,
        string sourceKind, OfficeDocumentRegion? region = null) {
        foreach (IWorkTextParagraph paragraph in content.Paragraphs) {
            AddParagraph(page, paragraph, sourceKind, region);
        }
    }

    private void AddTextBox(OfficeDocumentPage page, IWorkTextBox textBox,
        string sourceKind) {
        OfficeDocumentRegion? region = Region(textBox.Geometry);
        ReportUnsupportedRotation(page, textBox.Geometry, sourceKind);
        int firstBlockIndex = _blocks.Count;
        AddRichContent(page, textBox.Content, sourceKind, region);
        if (!textBox.Content.Paragraphs.Any(paragraph => paragraph.Runs.Any(run => run.Text.Length > 0))
            && !string.IsNullOrWhiteSpace(textBox.AccessibilityDescription)) {
            string description = textBox.AccessibilityDescription!;
            AddBlock(page, "text-box", description, EscapeMarkdown(description, _cancellationToken), null, null,
                markdownPart: (offset, length) => EscapeMarkdown(description.Substring(offset, length), _cancellationToken),
                sourceKind: sourceKind, region: region);
        }
        if (!string.IsNullOrWhiteSpace(textBox.Hyperlink)) {
            OfficeDocumentBlock? anchor = _blocks.Skip(firstBlockIndex)
                .FirstOrDefault(block => !string.IsNullOrWhiteSpace(block.Text));
            ReaderLocation linkLocation = anchor?.Location
                ?? (_blocks.Count > firstBlockIndex ? _blocks[firstBlockIndex].Location
                    : Location(page, sourceKind: sourceKind,
                        anchor: "iwork-shape-" + (_links.Count + 1).ToString("D6", CultureInfo.InvariantCulture)));
            AddLink(page, textBox.Hyperlink!, linkLocation, region: region);
        }
    }

    private static OfficeDocumentRegion? Region(IWorkGeometry? geometry) =>
        geometry == null ? null : new OfficeDocumentRegion {
            X = geometry.LeftPoints,
            Y = geometry.TopPoints,
            Width = geometry.WidthPoints,
            Height = geometry.HeightPoints
        };

    private void AddParagraph(OfficeDocumentPage page, IWorkTextParagraph paragraph,
        string sourceKind, OfficeDocumentRegion? region = null) {
        string text = ParagraphText(paragraph, _cancellationToken);
        ReportParagraphDetails(page, paragraph);
        if (text.Length == 0) {
            AddBlock(page, paragraph.ListLevel >= 0 ? "list-item" : "paragraph",
                string.Empty, paragraph.ListLevel >= 0 ? RichTextMarkdown(paragraph, _cancellationToken) : "\n",
                paragraph.ListLevel >= 0 ? paragraph.ListLevel + 1 : null,
                paragraph.ListLabel, sourceKind: sourceKind, region: region);
            return;
        }
        bool heading = paragraph.ListLevel < 0 && sourceKind == "title";
        string markdown = (heading ? "# " : string.Empty) + RichTextMarkdown(paragraph, _cancellationToken);
        string kind = paragraph.ListLevel >= 0 ? "list-item" :
            sourceKind == "title" ? "heading" : "paragraph";
        ReaderLocation blockLocation = AddBlock(page, kind, text, markdown,
            paragraph.ListLevel >= 0 ? paragraph.ListLevel + 1 : null,
            paragraph.ListLabel,
            markdownPart: (offset, length) =>
                (heading && offset == 0 ? "# " : string.Empty)
                + RichTextMarkdown(paragraph, offset, length, _cancellationToken),
            sourceKind: sourceKind, region: region);
        AddRunLinks(page, paragraph.Runs, blockLocation);
    }

    private void ReportParagraphDetails(OfficeDocumentPage page, IWorkTextParagraph paragraph) {
        if (!_reportedUnsupportedLayoutBreaks && paragraph.BreakKind is
            IWorkParagraphBreakKind.Page or IWorkParagraphBreakKind.Section
                or IWorkParagraphBreakKind.Layout) {
            _reportedUnsupportedLayoutBreaks = true;
            _diagnostics.Add(new OfficeDocumentDiagnostic {
                Category = OfficeDocumentDiagnosticCategory.Content,
                Code = "IWORK_READER_LAYOUT_BREAK_UNSUPPORTED",
                Message = "Reader text and Markdown do not represent explicit page, section, or layout breaks; the iWork source model retains them.",
                Source = "OfficeIMO.Reader.IWork",
                Location = Location(page)
            });
        }
        if (!_reportedUnsupportedTextStyles
            && (paragraph.ListFontName != null || HasUnrepresentedParagraphStyle(paragraph.Style)
                || HasUnrepresentedRunStyle(paragraph.Style.TextStyle)
                || paragraph.Runs.Any(run => HasUnrepresentedRunStyle(run.Style)))) {
            _reportedUnsupportedTextStyles = true;
            _diagnostics.Add(new OfficeDocumentDiagnostic {
                Category = OfficeDocumentDiagnosticCategory.Content,
                Code = "IWORK_READER_TEXT_STYLE_PARTIAL",
                Message = "Reader Markdown cannot represent all source paragraph, underline, font, or color formatting; the iWork source model retains those styles.",
                Source = "OfficeIMO.Reader.IWork",
                Location = Location(page)
            });
        }
        if (paragraph.ListLevel > MaximumMarkdownListLevel && !_reportedMarkdownListDepthLimit) {
            _reportedMarkdownListDepthLimit = true;
            _diagnostics.Add(new OfficeDocumentDiagnostic {
                Category = OfficeDocumentDiagnosticCategory.Limit,
                Code = "IWORK_READER_LIST_DEPTH_TRUNCATED",
                Message = "Markdown indentation is capped at 128 list levels; source list levels remain on the blocks.",
                Source = "OfficeIMO.Reader.IWork",
                Location = Location(page)
            });
        }
    }

    private void AddPlainText(OfficeDocumentPage page, string text, string sourceKind) {
        if (string.IsNullOrEmpty(text)) return;
        AddBlock(page, sourceKind, text, EscapeMarkdown(text, _cancellationToken), null, null,
            markdownPart: (offset, length) => EscapeMarkdown(text.Substring(offset, length), _cancellationToken));
    }

    private ReaderLocation AddBlock(OfficeDocumentPage page, string kind, string text,
        string markdown, int? level, string? marker,
        ReaderTable? table = null,
        Func<int, int, string>? markdownPart = null,
        string? sourceKind = null, OfficeDocumentRegion? region = null,
        bool splitMarkdownIndependently = false, int? tableIndex = null, string? a1Range = null) {
        _cancellationToken.ThrowIfCancellationRequested();
        string id = "iwork-b" + (_blocks.Count + 1).ToString("D6", CultureInfo.InvariantCulture);
        ReaderLocation location = Location(page, sourceKind: sourceKind ?? kind, anchor: id);
        location.TableIndex = tableIndex;
        location.A1Range = a1Range;
        var block = new OfficeDocumentBlock {
            Id = id,
            Kind = kind,
            Text = text,
            Level = level,
            Marker = marker,
            Location = location,
            Region = region
        };
        _blocks.Add(block);
        _pageBlocks[page].Add(block);
        int maxChars = Math.Max(1, _readerOptions.MaxChars);
        bool independentMarkdownChunks = splitMarkdownIndependently || markdown.Length > maxChars;
        int partIndex = 0;
        int textOffset = 0;
        int markdownOffset = 0;
        bool split = text.Length > maxChars
            || independentMarkdownChunks && markdown.Length > maxChars;
        while (textOffset < text.Length
            || independentMarkdownChunks && markdownOffset < markdown.Length
            || partIndex == 0) {
            _cancellationToken.ThrowIfCancellationRequested();
            int length = ScalarSafeChunkLength(text, textOffset, maxChars);
            string part = length == 0 ? string.Empty : text.Substring(textOffset, length);
            int markdownLength = independentMarkdownChunks
                ? ScalarSafeChunkLength(markdown, markdownOffset, maxChars) : 0;
            ReaderLocation chunkLocation = Location(page, _chunks.Count, sourceKind ?? kind, id);
            chunkLocation.TableIndex = tableIndex;
            chunkLocation.A1Range = a1Range;
            _chunks.Add(new ReaderChunk {
                Id = id + "-" + partIndex.ToString("D3", CultureInfo.InvariantCulture),
                Kind = ReaderInputKind.IWork,
                Location = chunkLocation,
                Text = part,
                Markdown = independentMarkdownChunks
                    ? markdownLength == 0 ? string.Empty
                        : markdown.Substring(markdownOffset, markdownLength)
                    : split
                    ? markdownPart?.Invoke(textOffset, length) ?? (partIndex == 0 ? markdown : string.Empty)
                    : markdown,
                ContinuesPreviousChunk = partIndex > 0,
                Tables = partIndex == 0 && table != null ? new[] { table } : null,
                Warnings = split ? new[] { "Content was split at ReaderOptions.MaxChars." } : null
            });
            textOffset += length;
            markdownOffset += markdownLength;
            partIndex++;
        }
        return location;
    }

    private void ReportUnsupportedRotation(OfficeDocumentPage page, IWorkGeometry? geometry,
        string sourceKind) {
        if (geometry == null || Math.Abs(geometry.RotationDegrees) < 0.000001d) return;
        _diagnostics.Add(new OfficeDocumentDiagnostic {
            Category = OfficeDocumentDiagnosticCategory.Content,
            Code = "IWORK_READER_ROTATION_UNSUPPORTED",
            Message = "A source drawable's rotation is omitted from Reader regions; the iWork source model retains it.",
            Source = "OfficeIMO.Reader.IWork",
            Location = Location(page, sourceKind: sourceKind)
        });
    }

    private static int ScalarSafeChunkLength(string value, int offset, int maximum) {
        int length = Math.Min(maximum, Math.Max(0, value.Length - offset));
        if (length > 0 && offset + length < value.Length
            && char.IsHighSurrogate(value[offset + length - 1])
            && char.IsLowSurrogate(value[offset + length])) {
            length = length == 1 ? 2 : length - 1;
        }
        return length;
    }

    private static bool HasUnrepresentedParagraphStyle(IWorkParagraphStyle style) =>
        style.Alignment is IWorkTextAlignment.Center or IWorkTextAlignment.Right
            or IWorkTextAlignment.Justified
        || style.FirstLineIndentPoints is not null and not 0d
        || style.LeftIndentPoints is not null and not 0d
        || style.RightIndentPoints is not null and not 0d
        || style.SpaceBeforePoints is not null and not 0d
        || style.SpaceAfterPoints is not null and not 0d
        || style.LineSpacingMultiplier.HasValue
        || style.TabStops?.Count > 0
        || style.PageBreakBefore == true || style.KeepWithNext == true
        || style.KeepLinesTogether == true;

    private static bool HasUnrepresentedRunStyle(IWorkTextStyle style) =>
        style.Underline == true || style.FontSizePoints.HasValue
        || style.FontName != null || style.Color != null
        || style.BackgroundColor != null;

    private ReaderLocation Location(OfficeDocumentPage page, int? blockIndex = null,
        string? sourceKind = null, string? anchor = null) => new ReaderLocation {
            Path = _path,
            Sheet = page.Location.Sheet,
            Slide = page.Location.Slide,
            BlockIndex = blockIndex,
            SourceBlockKind = sourceKind,
            BlockAnchor = anchor
        };

    private void AddDiagnostics(IEnumerable<IWorkDiagnostic> diagnostics) {
        foreach (IWorkDiagnostic diagnostic in diagnostics) {
            _diagnostics.Add(new OfficeDocumentDiagnostic {
                Severity = diagnostic.Severity switch {
                    IWorkDiagnosticSeverity.Error => OfficeDocumentDiagnosticSeverity.Error,
                    IWorkDiagnosticSeverity.Warning => OfficeDocumentDiagnosticSeverity.Warning,
                    _ => OfficeDocumentDiagnosticSeverity.Information
                },
                Category = diagnostic.Severity == IWorkDiagnosticSeverity.Information
                    ? OfficeDocumentDiagnosticCategory.Adapter
                    : OfficeDocumentDiagnosticCategory.Content,
                Code = diagnostic.Code,
                Message = diagnostic.Message,
                Source = "OfficeIMO.IWork",
                Location = new ReaderLocation {
                    Path = diagnostic.EntryPath == null
                        ? _path : _path + "!/" + diagnostic.EntryPath,
                    SourceBlockKind = diagnostic.RecordIdentifier.HasValue
                        ? "iwa-record" : null
                },
                Attributes = DiagnosticAttributes(diagnostic)
            });
        }
    }

    private static IReadOnlyDictionary<string, string> DiagnosticAttributes(
        IWorkDiagnostic diagnostic) {
        var attributes = new Dictionary<string, string>(StringComparer.Ordinal) {
            ["lossKind"] = diagnostic.LossKind.ToString()
        };
        if (diagnostic.EntryPath != null) attributes.Add("entryPath", diagnostic.EntryPath);
        if (diagnostic.RecordIdentifier.HasValue) {
            attributes.Add("recordIdentifier", diagnostic.RecordIdentifier.Value.ToString(
                CultureInfo.InvariantCulture));
        }
        return attributes;
    }
}
