using OfficeIMO.IWork;

namespace OfficeIMO.Reader.IWork;

internal sealed partial class IWorkReadProjection {
    private void AddTable(OfficeDocumentPage page, IWorkTable source) {
        _cancellationToken.ThrowIfCancellationRequested();
        int tableIndex = _pageTables[page].Count;
        int rowCount = Math.Min(source.RowCount, Math.Max(1, _readerOptions.MaxTableRows));
        int columnCount = Math.Min(source.ColumnCount, _options.MaximumTableColumns);
        if (rowCount == 0 || columnCount == 0) return;
        bool hasHeader = source.HeaderRowCount > 0;
        string[] columns = Enumerable.Range(1, columnCount)
            .Select(column => hasHeader
                ? CellText(source.GetCell(1, column))
                : "Column " + column.ToString(CultureInfo.InvariantCulture))
            .ToArray();
        var rows = new List<IReadOnlyList<string>>();
        for (int row = hasHeader ? 2 : 1; row <= rowCount; row++) {
            _cancellationToken.ThrowIfCancellationRequested();
            rows.Add(Enumerable.Range(1, columnCount)
                .Select(column => CellText(source.GetCell(row, column))).ToArray());
        }
        bool truncated = source.RowCount > rowCount || source.ColumnCount > columnCount;
        var location = Location(page);
        location.TableIndex = tableIndex;
        var table = new ReaderTable {
            Title = source.Name,
            Kind = "iwork-table",
            Location = location,
            Columns = columns,
            Rows = rows,
            TotalRowCount = source.RowCount,
            Truncated = truncated
        };
        _tables.Add(table);
        _pageTables[page].Add(table);
        string markdown = table.ToMarkdownTable();
        AddBlock(page, "table", markdown, markdown, null, null, table);
        if (truncated) {
            _diagnostics.Add(new OfficeDocumentDiagnostic {
                Category = OfficeDocumentDiagnosticCategory.Limit,
                Code = "IWORK_READER_TABLE_TRUNCATED",
                Message = $"Table '{source.Name}' exceeds Reader row or column materialization limits.",
                Source = "OfficeIMO.Reader.IWork",
                Location = location
            });
        }
        if (source.Cells.Any(cell => cell.Kind == IWorkCellKind.Formula && cell.Value != null)) {
            _diagnostics.Add(new OfficeDocumentDiagnostic {
                Category = OfficeDocumentDiagnosticCategory.Content,
                Code = "IWORK_READER_FORMULA_CACHE",
                Message = $"Table '{source.Name}' presents formula cached values; formulas remain on the iWork source model.",
                Source = "OfficeIMO.Reader.IWork",
                Location = location
            });
        }
        foreach (IWorkTableCell cell in source.Cells) {
            if (cell.Row > rowCount || cell.Column > columnCount || cell.RichText == null) continue;
            foreach (IWorkTextParagraph paragraph in cell.RichText.Paragraphs) {
                AddRunLinks(page, paragraph.Runs, location);
            }
        }
    }

    private static string CellText(IWorkTableCell? cell) => cell == null
        ? string.Empty
        : cell.Kind == IWorkCellKind.Formula && cell.Value != null
            ? cell.CachedDisplayText
            : cell.DisplayText;

    private void AddImage(OfficeDocumentPage page, IWorkImageAsset source) {
        _cancellationToken.ThrowIfCancellationRequested();
        string id = "iwork-a" + (_assets.Count + 1).ToString("D6", CultureInfo.InvariantCulture);
        var asset = new OfficeDocumentAsset {
            Id = id,
            Kind = "image",
            MediaType = source.MediaType,
            FileName = source.FileName,
            Extension = Path.GetExtension(source.FileName),
            AltText = source.AccessibilityDescription,
            Width = source.PixelWidth,
            Height = source.PixelHeight,
            LengthBytes = source.Length,
            SourceObjectId = source.PackagePath,
            PayloadBytes = _options.IncludeImagePayloads ? source.GetBytes() : null,
            Location = Location(page),
            Region = source.Geometry == null ? null : new OfficeDocumentRegion {
                X = source.Geometry.LeftPoints,
                Y = source.Geometry.TopPoints,
                Width = source.Geometry.WidthPoints,
                Height = source.Geometry.HeightPoints
            }
        };
        _assets.Add(asset);
        _pageAssets[page].Add(asset);
        if (!string.IsNullOrWhiteSpace(source.AccessibilityDescription)) {
            AddBlock(page, "image", source.AccessibilityDescription!,
                source.AccessibilityDescription!, null, null);
        }
        if (source.Hyperlink != null) AddLink(page, source.Hyperlink, asset.Location);
        if (source.HasMask) {
            _diagnostics.Add(new OfficeDocumentDiagnostic {
                Category = OfficeDocumentDiagnosticCategory.Content,
                Code = "IWORK_READER_IMAGE_MASK",
                Message = "The source image has a mask or crop that the Reader asset does not apply.",
                Source = "OfficeIMO.Reader.IWork",
                Location = asset.Location
            });
        }
    }

    private void AddRunLinks(OfficeDocumentPage page,
        IEnumerable<IWorkTextRun> runs, ReaderLocation? location = null) {
        foreach (IWorkTextRun run in runs) {
            if (run.Hyperlink != null) AddLink(page, run.Hyperlink, location ?? Location(page));
        }
    }

    private void AddLink(OfficeDocumentPage page, string target, ReaderLocation location) {
        var link = new OfficeDocumentLink {
            Id = "iwork-l" + (_links.Count + 1).ToString("D6", CultureInfo.InvariantCulture),
            Kind = "uri",
            Uri = target,
            Location = location
        };
        _links.Add(link);
        _pageLinks[page].Add(link);
    }

    private static string RichTextMarkdown(IWorkTextParagraph paragraph) {
        var builder = new StringBuilder();
        foreach (IWorkTextRun run in paragraph.Runs) {
            string value = EscapeMarkdown(run.Text);
            if (run.Style.Bold == true) value = "**" + value + "**";
            if (run.Style.Italic == true) value = "*" + value + "*";
            if (run.Hyperlink != null
                && Uri.TryCreate(run.Hyperlink, UriKind.Absolute, out Uri? uri)) {
                value = "[" + value + "](<" + uri.AbsoluteUri.Replace(">", "%3E") + ">)";
            }
            builder.Append(value);
        }
        if (paragraph.ListLevel >= 0) {
            return (string.IsNullOrWhiteSpace(paragraph.ListLabel) ? "-" : paragraph.ListLabel)
                + " " + builder;
        }
        return builder.ToString();
    }

    private static string EscapeMarkdown(string value) => value
        .Replace("\\", "\\\\")
        .Replace("[", "\\[")
        .Replace("]", "\\]")
        .Replace("*", "\\*");
}
