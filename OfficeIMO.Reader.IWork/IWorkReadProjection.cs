using OfficeIMO.IWork;

namespace OfficeIMO.Reader.IWork;

internal sealed partial class IWorkReadProjection {
    private readonly OfficeDocumentReadResult _result;
    private readonly string _path;
    private readonly ReaderOptions _readerOptions;
    private readonly ReaderIWorkOptions _options;
    private readonly CancellationToken _cancellationToken;
    private readonly List<ReaderChunk> _chunks = new();
    private readonly List<OfficeDocumentBlock> _blocks = new();
    private readonly List<ReaderTable> _tables = new();
    private readonly List<OfficeDocumentAsset> _assets = new();
    private readonly List<OfficeDocumentLink> _links = new();
    private readonly List<OfficeDocumentPage> _pages = new();
    private readonly List<OfficeDocumentDiagnostic> _diagnostics = new();
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
    }

    internal void AddPages(IWorkPagesProjection source) {
        var page = NewPage(null, "Pages document", null);
        foreach (IWorkTextParagraph paragraph in source.Body.Paragraphs) {
            AddParagraph(page, paragraph, "body");
        }
        foreach (IWorkPagesDrawable drawable in source.Drawables) {
            _cancellationToken.ThrowIfCancellationRequested();
            switch (drawable.Kind) {
                case IWorkPagesDrawableKind.TextBox:
                    AddTextBox(page, drawable.TextBox!, "text-box");
                    break;
                case IWorkPagesDrawableKind.Image:
                    AddImage(page, drawable.Image!);
                    break;
                case IWorkPagesDrawableKind.Table:
                    AddTable(page, drawable.Table!);
                    break;
            }
        }
        foreach (IWorkTextContent header in source.HeaderContents) AddRichContent(page, header, "header");
        foreach (IWorkTextContent footer in source.FooterContents) AddRichContent(page, footer, "footer");
        if (source.PageLayout is { } layout) {
            page.Width = layout.WidthPoints;
            page.Height = layout.HeightPoints;
        }
        AddDiagnostics(source.Diagnostics);
    }

    internal void AddNumbers(IWorkNumbersProjection source) {
        for (int sheetIndex = 0; sheetIndex < source.Sheets.Count; sheetIndex++) {
            _cancellationToken.ThrowIfCancellationRequested();
            IWorkNumbersSheet sheet = source.Sheets[sheetIndex];
            var page = NewPage(sheetIndex + 1, sheet.Name, sheet.Name);
            foreach (string textBox in sheet.TextBoxes) AddPlainText(page, textBox, "text-box");
            foreach (IWorkTable table in sheet.Tables) AddTable(page, table);
        }
        AddDiagnostics(source.Diagnostics);
    }

    internal void AddKeynote(IWorkKeynoteProjection source) {
        foreach (IWorkKeynoteSlide slide in source.Slides) {
            _cancellationToken.ThrowIfCancellationRequested();
            var page = NewPage(slide.Index, slide.Name, null, slide.Index);
            page.Width = source.SlideSize?.WidthPoints;
            page.Height = source.SlideSize?.HeightPoints;
            foreach (IWorkKeynoteDrawable drawable in slide.Drawables) {
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
        _result.Markdown = string.Join("\n\n", _chunks.Select(chunk => chunk.Markdown)
            .Where(markdown => !string.IsNullOrEmpty(markdown)));
        _result.Metadata = source.BuildVersions.Select((version, index) =>
            new OfficeDocumentMetadataEntry {
                Id = "iwork-build-" + index.ToString("D4", CultureInfo.InvariantCulture),
                Category = "producer",
                Name = "BuildVersion",
                Value = version
            }).ToArray();
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
        string sourceKind) {
        foreach (IWorkTextParagraph paragraph in content.Paragraphs) {
            AddParagraph(page, paragraph, sourceKind);
        }
    }

    private void AddTextBox(OfficeDocumentPage page, IWorkTextBox textBox,
        string sourceKind) {
        AddRichContent(page, textBox.Content, sourceKind);
        if (!string.IsNullOrWhiteSpace(textBox.Hyperlink)) {
            AddLink(page, textBox.Hyperlink!, Location(page));
        }
    }

    private void AddParagraph(OfficeDocumentPage page, IWorkTextParagraph paragraph,
        string sourceKind) {
        string text = paragraph.Text;
        if (text.Length == 0) return;
        string markdown = RichTextMarkdown(paragraph);
        string kind = paragraph.ListLevel >= 0 ? "list-item" :
            sourceKind == "title" ? "heading" : "paragraph";
        AddBlock(page, kind, text, markdown,
            paragraph.ListLevel >= 0 ? paragraph.ListLevel + 1 : null,
            paragraph.ListLabel);
        AddRunLinks(page, paragraph.Runs);
    }

    private void AddPlainText(OfficeDocumentPage page, string text, string sourceKind) {
        if (string.IsNullOrEmpty(text)) return;
        AddBlock(page, sourceKind, text, text, null, null);
    }

    private void AddBlock(OfficeDocumentPage page, string kind, string text,
        string markdown, int? level, string? marker,
        ReaderTable? table = null) {
        _cancellationToken.ThrowIfCancellationRequested();
        string id = "iwork-b" + (_blocks.Count + 1).ToString("D6", CultureInfo.InvariantCulture);
        ReaderLocation location = Location(page);
        var block = new OfficeDocumentBlock {
            Id = id,
            Kind = kind,
            Text = text,
            Level = level,
            Marker = marker,
            Location = location
        };
        _blocks.Add(block);
        _pageBlocks[page].Add(block);
        int maxChars = Math.Max(1, _readerOptions.MaxChars);
        int partIndex = 0;
        for (int offset = 0; offset < text.Length || offset == 0; offset += maxChars) {
            int length = Math.Min(maxChars, text.Length - offset);
            bool split = text.Length > maxChars;
            string part = text.Substring(offset, length);
            _chunks.Add(new ReaderChunk {
                Id = id + "-" + partIndex.ToString("D3", CultureInfo.InvariantCulture),
                Kind = ReaderInputKind.IWork,
                Location = Location(page, _chunks.Count, kind, id),
                Text = part,
                Markdown = split ? part : markdown,
                Tables = partIndex == 0 && table != null ? new[] { table } : null,
                Warnings = split ? new[] { "Content was split at ReaderOptions.MaxChars." } : null
            });
            partIndex++;
            if (text.Length == 0) break;
        }
    }

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
                Location = new ReaderLocation { Path = _path }
            });
        }
    }
}
