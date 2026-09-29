using OfficeIMO.IWork;

namespace OfficeIMO.Reader.IWork;

internal sealed partial class IWorkReadProjection {
    private void AddTable(OfficeDocumentPage page, IWorkTable source) {
        _cancellationToken.ThrowIfCancellationRequested();
        ReportUnsupportedRotation(page, source.Geometry, "table");
        int tableIndex = _pageTables[page].Count;
        if (source.RowCount == 0 || source.ColumnCount == 0) return;
        int remainingCells = _options.MaximumProjectedTableCells - _projectedTableCells;
        if (remainingCells == 0) {
            if (!_reportedTableBudgetExhausted) {
                _reportedTableBudgetExhausted = true;
                _diagnostics.Add(new OfficeDocumentDiagnostic {
                    Category = OfficeDocumentDiagnosticCategory.Limit,
                    Code = "IWORK_READER_TABLE_BUDGET_EXCEEDED",
                    Message = "Additional iWork tables were omitted because the Reader document-wide dense table cell limit was reached.",
                    Source = "OfficeIMO.Reader.IWork",
                    Location = Location(page)
                });
            }
            return;
        }
        int columnCount = Math.Min(Math.Min(source.ColumnCount, _options.MaximumTableColumns),
            remainingCells);
        int remainingRows = remainingCells / columnCount;
        int headerRows = Math.Min(source.HeaderRowCount, source.RowCount);
        int materializedHeaderRows = Math.Min(headerRows,
            Math.Min(Math.Max(1, _readerOptions.MaxTableRows), remainingRows));
        int totalDataRows = source.RowCount - headerRows;
        int dataRows = Math.Min(totalDataRows,
            Math.Min(Math.Max(1, _readerOptions.MaxTableRows),
                remainingRows - materializedHeaderRows));
        _projectedTableCells += columnCount * (materializedHeaderRows + dataRows);
        bool hasHeader = headerRows > 0;
        string[] columns = Enumerable.Range(1, columnCount)
            .Select(column => hasHeader
                ? string.Join(" / ", Enumerable.Range(1, materializedHeaderRows)
                    .Select(row => CellText(source.GetCell(row, column)))
                    .Where(value => value.Length > 0))
                : "Column " + column.ToString(CultureInfo.InvariantCulture))
            .ToArray();
        var rows = new List<IReadOnlyList<string>>();
        for (int row = headerRows + 1; row <= headerRows + dataRows; row++) {
            _cancellationToken.ThrowIfCancellationRequested();
            rows.Add(Enumerable.Range(1, columnCount)
                .Select(column => CellText(source.GetCell(row, column))).ToArray());
        }
        bool truncated = headerRows > materializedHeaderRows
            || totalDataRows > dataRows || source.ColumnCount > columnCount;
        _cancellationToken.ThrowIfCancellationRequested();
        IReadOnlyList<ReaderTableColumnProfile> profiles = ReaderTableProfiler.CreateProfiles(columns, rows);
        _cancellationToken.ThrowIfCancellationRequested();
        var location = Location(page);
        location.TableIndex = tableIndex;
        var table = new ReaderTable {
            Title = source.Name,
            Summary = source.AccessibilityDescription,
            Kind = "iwork-table",
            Location = location,
            Columns = columns,
            ColumnProfiles = profiles,
            Diagnostics = TableDiagnostics(source),
            Rows = rows,
            TotalRowCount = totalDataRows,
            Truncated = truncated
        };
        _tables.Add(table);
        _pageTables[page].Add(table);
        string markdown = table.ToMarkdownTable(_cancellationToken);
        string text = DocumentReaderEngine.BuildRichTableText(table, _cancellationToken);
        ReaderLocation tableBlockLocation = AddBlock(page, "table", text, markdown, null, null, table,
            region: Region(source.Geometry), splitMarkdownIndependently: true,
            tableIndex: tableIndex);
        if (truncated) {
            _diagnostics.Add(new OfficeDocumentDiagnostic {
                Category = OfficeDocumentDiagnosticCategory.Limit,
                Code = "IWORK_READER_TABLE_TRUNCATED",
                Message = $"Table '{source.Name}' exceeds Reader row, column, or document-wide dense cell materialization limits.",
                Source = "OfficeIMO.Reader.IWork",
                Location = location
            });
        }
        if (source.MergedRanges.Count > 0) {
            _diagnostics.Add(new OfficeDocumentDiagnostic {
                Category = OfficeDocumentDiagnosticCategory.Content,
                Code = "IWORK_READER_TABLE_MERGES_UNSUPPORTED",
                Message = $"Table '{source.Name}' is projected as a flat Reader grid; merged ranges remain on the iWork source model.",
                Source = "OfficeIMO.Reader.IWork",
                Location = location
            });
        }
        if (source.HeaderColumnCount > 0 || source.FooterRowCount > 0) {
            _diagnostics.Add(new OfficeDocumentDiagnostic {
                Category = OfficeDocumentDiagnosticCategory.Content,
                Code = "IWORK_READER_TABLE_ROLES_UNSUPPORTED",
                Message = $"Table '{source.Name}' is projected as a flat Reader grid; source header-column and footer-row roles remain on the iWork source model.",
                Source = "OfficeIMO.Reader.IWork",
                Location = location,
                Attributes = new Dictionary<string, string>(StringComparer.Ordinal) {
                    ["headerColumnCount"] = source.HeaderColumnCount.ToString(CultureInfo.InvariantCulture),
                    ["footerRowCount"] = source.FooterRowCount.ToString(CultureInfo.InvariantCulture)
                }
            });
        }
        long sourceCellArea = (long)source.RowCount * source.ColumnCount;
        if (source.Geometry != null && sourceCellArea > int.MaxValue) {
            _diagnostics.Add(new OfficeDocumentDiagnostic {
                Category = OfficeDocumentDiagnosticCategory.Limit,
                Code = "IWORK_READER_TABLE_CELL_COUNTS_SATURATED",
                Message = $"Table '{source.Name}' has more logical cells than Reader's 32-bit table diagnostic counts can represent; expected and missing counts are saturated.",
                Source = "OfficeIMO.Reader.IWork",
                Location = location,
                Attributes = new Dictionary<string, string>(StringComparer.Ordinal) {
                    ["sourceLogicalCellCount"] = sourceCellArea.ToString(CultureInfo.InvariantCulture)
                }
            });
        }
        bool hasFormula = false;
        bool hasUnrepresentedStyle = false;
        foreach (IWorkTableCell cell in source.Cells) {
            _cancellationToken.ThrowIfCancellationRequested();
            hasFormula |= cell.Kind == IWorkCellKind.Formula;
            if (cell.RichText is not { } richText) continue;
            hasUnrepresentedStyle |= !richText.IsComplete;
            bool isProjected = cell.Row <= headerRows + dataRows && cell.Column <= columnCount
                && (cell.Row <= materializedHeaderRows || cell.Row > headerRows);
            foreach (IWorkTextParagraph paragraph in richText.Paragraphs) {
                _cancellationToken.ThrowIfCancellationRequested();
                hasUnrepresentedStyle |= HasUnrepresentedParagraphStyle(paragraph.Style)
                    || HasUnrepresentedRunStyle(paragraph.Style.TextStyle);
                foreach (IWorkTextRun run in paragraph.Runs) {
                    _cancellationToken.ThrowIfCancellationRequested();
                    hasUnrepresentedStyle |= run.Style.Bold == true
                        || run.Style.Italic == true || run.Style.Underline == true
                        || run.Style.Strikethrough == true || run.Style.FontSizePoints.HasValue
                        || run.Style.FontName != null || run.Style.Color != null
                        || run.Style.BackgroundColor != null;
                    if (isProjected && run.Hyperlink != null)
                        AddLink(page, run.Hyperlink, tableBlockLocation, run.Text);
                }
            }
        }
        if (hasFormula) {
            _diagnostics.Add(new OfficeDocumentDiagnostic {
                Category = OfficeDocumentDiagnosticCategory.Content,
                Code = "IWORK_READER_FORMULA_CACHE",
                Message = $"Table '{source.Name}' presents formula cached values; formulas remain on the iWork source model.",
                Source = "OfficeIMO.Reader.IWork",
                Location = location
            });
        }
        if (hasUnrepresentedStyle) {
            _diagnostics.Add(new OfficeDocumentDiagnostic {
                Category = OfficeDocumentDiagnosticCategory.Content,
                Code = "IWORK_READER_TABLE_STYLE_PARTIAL",
                Message = $"Table '{source.Name}' is projected as plain Reader table text; source rich-text cell formatting is unresolved or cannot be represented in Reader output.",
                Source = "OfficeIMO.Reader.IWork",
                Location = location
            });
        }
    }

    private static ReaderTableDiagnostics? TableDiagnostics(IWorkTable source) {
        if (source.Geometry is not { } geometry) return null;
        long sourceArea = (long)source.RowCount * source.ColumnCount;
        int expectedCells = (int)Math.Min(sourceArea, int.MaxValue);
        int filledCells = source.Cells.Count;
        int missingCells = (int)Math.Min(Math.Max(0, sourceArea - filledCells), int.MaxValue);
        return new ReaderTableDiagnostics {
            Confidence = 1,
            SchemaConfidence = 1,
            CellCompleteness = sourceArea == 0 ? 1 : (double)filledCells / sourceArea,
            ColumnGeometryConfidence = source.DefaultColumnWidth.HasValue ? 1 : 0,
            SourceRowCount = source.RowCount,
            ExpectedCellCount = expectedCells,
            FilledCellCount = filledCells,
            MissingCellCount = missingCells,
            HasGeometry = true,
            XStart = geometry.LeftPoints,
            XEnd = geometry.LeftPoints + geometry.WidthPoints,
            YTop = geometry.TopPoints,
            YBottom = geometry.TopPoints + geometry.HeightPoints,
            Width = geometry.WidthPoints,
            Height = geometry.HeightPoints
        };
    }

    private static string CellText(IWorkTableCell? cell) => cell == null
        ? string.Empty
        : cell.Kind == IWorkCellKind.Formula
            ? cell.CachedDisplayText
            : cell.DisplayText;

    private void AddImage(OfficeDocumentPage page, IWorkImageAsset source) {
        _cancellationToken.ThrowIfCancellationRequested();
        ReportUnsupportedRotation(page, source.Geometry, "image");
        string id = "iwork-a" + (_assets.Count + 1).ToString("D6", CultureInfo.InvariantCulture);
        var asset = new OfficeDocumentAsset {
            Id = id,
            Kind = "image",
            MediaType = source.MediaType,
            FileName = OfficeDocumentAssetNaming.BuildFileName(id,
                Path.GetExtension(source.FileName)),
            Extension = Path.GetExtension(source.FileName),
            AltText = source.AccessibilityDescription,
            Width = source.PixelWidth,
            Height = source.PixelHeight,
            LengthBytes = source.Length,
            SourceObjectId = source.PackagePath,
            PayloadBytes = _options.IncludeImagePayloads ? source.GetBytes() : null,
            Location = Location(page, sourceKind: "image", anchor: id),
            Region = Region(source.Geometry)
        };
        _assets.Add(asset);
        _pageAssets[page].Add(asset);
        if (!string.IsNullOrWhiteSpace(source.AccessibilityDescription)) {
            string description = source.AccessibilityDescription!;
            AddBlock(page, "image", description, EscapeMarkdown(description, _cancellationToken), null, null,
                markdownPart: (offset, length) => EscapeMarkdown(description.Substring(offset, length), _cancellationToken),
                region: asset.Region);
        }
        if (source.Hyperlink != null) AddLink(page, source.Hyperlink, asset.Location,
            region: asset.Region);
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
            _cancellationToken.ThrowIfCancellationRequested();
            if (run.Hyperlink != null) AddLink(page, run.Hyperlink, location ?? Location(page), run.Text);
        }
    }

    private void AddLink(OfficeDocumentPage page, string target, ReaderLocation location,
        string? text = null, OfficeDocumentRegion? region = null) {
        var link = new OfficeDocumentLink {
            Id = "iwork-l" + (_links.Count + 1).ToString("D6", CultureInfo.InvariantCulture),
            Kind = "uri",
            Uri = target,
            Text = text,
            Location = location,
            Region = region
        };
        _links.Add(link);
        _pageLinks[page].Add(link);
    }

    internal static string RichTextMarkdown(IWorkTextParagraph paragraph,
        CancellationToken cancellationToken = default) =>
        RichTextMarkdown(paragraph, 0, int.MaxValue, cancellationToken);

    private static string ParagraphText(IWorkTextParagraph paragraph,
        CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        var builder = new StringBuilder();
        foreach (IWorkTextRun run in paragraph.Runs) {
            string text = run.Text;
            for (int offset = 0; offset < text.Length; offset += 4096) {
                cancellationToken.ThrowIfCancellationRequested();
                builder.Append(text, offset, Math.Min(4096, text.Length - offset));
            }
        }
        cancellationToken.ThrowIfCancellationRequested();
        return builder.ToString();
    }

    private static string RichTextMarkdown(IWorkTextParagraph paragraph, int offset, int length,
        CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        var builder = new StringBuilder();
        int runOffset = 0;
        foreach (IWorkTextRun run in paragraph.Runs) {
            cancellationToken.ThrowIfCancellationRequested();
            int runStart = runOffset;
            runOffset += run.Text.Length;
            int start = Math.Max(offset, runStart);
            int end = Math.Min(offset + length, runOffset);
            if (end <= start) continue;
            string segment = run.Text.Substring(start - runStart, end - start);
            int leading = 0;
            while (leading < segment.Length && char.IsWhiteSpace(segment[leading])) {
                if ((leading & 4095) == 0) cancellationToken.ThrowIfCancellationRequested();
                leading++;
            }
            int trailing = segment.Length;
            while (trailing > leading && char.IsWhiteSpace(segment[trailing - 1])) {
                if ((trailing & 4095) == 0) cancellationToken.ThrowIfCancellationRequested();
                trailing--;
            }
            string value = EscapeMarkdown(segment.Substring(leading, trailing - leading), cancellationToken);
            if (value.Length > 0) {
                if (run.Style.Bold == true) value = "**" + value + "**";
                if (run.Style.Italic == true) value = "*" + value + "*";
                if (run.Style.Strikethrough == true) value = "~~" + value + "~~";
            }
            value = EscapeMarkdown(segment.Substring(0, leading), cancellationToken) + value
                + EscapeMarkdown(segment.Substring(trailing), cancellationToken);
            if (run.Hyperlink != null
                && Uri.TryCreate(run.Hyperlink, UriKind.Absolute, out Uri? uri)
                && (uri.Scheme == Uri.UriSchemeHttp || uri.Scheme == Uri.UriSchemeHttps
                    || uri.Scheme == Uri.UriSchemeMailto)) {
                value = "[" + value + "](<" + uri.AbsoluteUri.Replace(">", "%3E") + ">)";
            }
            cancellationToken.ThrowIfCancellationRequested();
            builder.Append(value);
        }
        if (offset == 0 && paragraph.ListLevel >= 0) {
            string marker = MarkdownListMarker(paragraph.ListLabel);
            return new string(' ', Math.Min(paragraph.ListLevel, MaximumMarkdownListLevel) * 2)
                + marker + " " + builder;
        }
        return builder.ToString();
    }

    private static string MarkdownListMarker(string? sourceLabel) {
        string label = sourceLabel?.Trim() ?? string.Empty;
        int digits = 0;
        while (digits < label.Length && label[digits] >= '0' && label[digits] <= '9') digits++;
        return digits > 0 && digits == label.Length - 1
            && (label[digits] == '.' || label[digits] == ')')
            ? label.Substring(0, digits) + "."
            : "-";
    }

    private static string EscapeMarkdown(string value, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        var builder = new StringBuilder(value.Length);
        for (int index = 0; index < value.Length; index++) {
            if ((index & 4095) == 0) cancellationToken.ThrowIfCancellationRequested();
            char character = value[index];
            if ("\\`*_{}[]()#+-.!>|~".IndexOf(character) >= 0) builder.Append('\\');
            builder.Append(character);
        }
        cancellationToken.ThrowIfCancellationRequested();
        return builder.ToString();
    }
}
