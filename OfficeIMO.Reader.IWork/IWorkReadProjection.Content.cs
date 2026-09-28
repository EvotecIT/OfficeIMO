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
        var location = Location(page);
        location.TableIndex = tableIndex;
        var table = new ReaderTable {
            Title = source.Name,
            Summary = source.AccessibilityDescription,
            Kind = "iwork-table",
            Location = location,
            Columns = columns,
            Rows = rows,
            TotalRowCount = totalDataRows,
            Truncated = truncated
        };
        _tables.Add(table);
        _pageTables[page].Add(table);
        string markdown = table.ToMarkdownTable();
        string text = DocumentReaderEngine.BuildRichTableText(table);
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
        if (source.Cells.Any(cell => cell.Kind == IWorkCellKind.Formula)) {
            _diagnostics.Add(new OfficeDocumentDiagnostic {
                Category = OfficeDocumentDiagnosticCategory.Content,
                Code = "IWORK_READER_FORMULA_CACHE",
                Message = $"Table '{source.Name}' presents formula cached values; formulas remain on the iWork source model.",
                Source = "OfficeIMO.Reader.IWork",
                Location = location
            });
        }
        if (source.Cells.Any(cell => cell.RichText is { } richText
                && (!richText.IsComplete || richText.Paragraphs.Any(paragraph =>
                    HasUnrepresentedParagraphStyle(paragraph.Style)
                    || HasUnrepresentedRunStyle(paragraph.Style.TextStyle)
                    || paragraph.Runs.Any(run => run.Style.Bold == true
                        || run.Style.Italic == true || run.Style.Underline == true
                        || run.Style.Strikethrough == true || run.Style.FontSizePoints.HasValue
                        || run.Style.FontName != null || run.Style.Color != null
                        || run.Style.BackgroundColor != null))))) {
            _diagnostics.Add(new OfficeDocumentDiagnostic {
                Category = OfficeDocumentDiagnosticCategory.Content,
                Code = "IWORK_READER_TABLE_STYLE_PARTIAL",
                Message = $"Table '{source.Name}' is projected as plain Reader table text; source rich-text cell formatting is unresolved or cannot be represented in Reader output.",
                Source = "OfficeIMO.Reader.IWork",
                Location = location
            });
        }
        foreach (IWorkTableCell cell in source.Cells) {
            if (cell.Row > headerRows + dataRows || cell.Column > columnCount
                || (cell.Row > materializedHeaderRows && cell.Row <= headerRows)
                || cell.RichText == null) continue;
            foreach (IWorkTextParagraph paragraph in cell.RichText.Paragraphs) {
                AddRunLinks(page, paragraph.Runs, tableBlockLocation);
            }
        }
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
            AddBlock(page, "image", description, EscapeMarkdown(description), null, null,
                markdownPart: (offset, length) => EscapeMarkdown(description.Substring(offset, length)),
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

    internal static string RichTextMarkdown(IWorkTextParagraph paragraph) =>
        RichTextMarkdown(paragraph, 0, paragraph.Text.Length);

    private static string RichTextMarkdown(IWorkTextParagraph paragraph, int offset, int length) {
        var builder = new StringBuilder();
        int runOffset = 0;
        foreach (IWorkTextRun run in paragraph.Runs) {
            int runStart = runOffset;
            runOffset += run.Text.Length;
            int start = Math.Max(offset, runStart);
            int end = Math.Min(offset + length, runOffset);
            if (end <= start) continue;
            string segment = run.Text.Substring(start - runStart, end - start);
            int leading = 0;
            while (leading < segment.Length && char.IsWhiteSpace(segment[leading])) leading++;
            int trailing = segment.Length;
            while (trailing > leading && char.IsWhiteSpace(segment[trailing - 1])) trailing--;
            string value = EscapeMarkdown(segment.Substring(leading, trailing - leading));
            if (value.Length > 0) {
                if (run.Style.Bold == true) value = "**" + value + "**";
                if (run.Style.Italic == true) value = "*" + value + "*";
                if (run.Style.Strikethrough == true) value = "~~" + value + "~~";
            }
            value = EscapeMarkdown(segment.Substring(0, leading)) + value
                + EscapeMarkdown(segment.Substring(trailing));
            if (run.Hyperlink != null
                && Uri.TryCreate(run.Hyperlink, UriKind.Absolute, out Uri? uri)
                && (uri.Scheme == Uri.UriSchemeHttp || uri.Scheme == Uri.UriSchemeHttps
                    || uri.Scheme == Uri.UriSchemeMailto)) {
                value = "[" + value + "](<" + uri.AbsoluteUri.Replace(">", "%3E") + ">)";
            }
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

    private static string EscapeMarkdown(string value) {
        var builder = new StringBuilder(value.Length);
        foreach (char character in value) {
            if ("\\`*_{}[]()#+-.!>|~".IndexOf(character) >= 0) builder.Append('\\');
            builder.Append(character);
        }
        return builder.ToString();
    }
}
