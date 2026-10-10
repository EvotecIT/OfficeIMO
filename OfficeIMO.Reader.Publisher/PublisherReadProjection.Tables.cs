using OfficeIMO.Publisher;

namespace OfficeIMO.Reader.Publisher;

internal sealed partial class PublisherReadProjection {
    private readonly List<ReaderTable> _tables = new();
    private readonly Dictionary<uint, List<ReaderTable>> _storyTables = new();
    private readonly Dictionary<uint, List<ReaderTable>> _pageTables = new();
    private readonly Dictionary<uint, List<OfficeDocumentBlock>> _pageTableBlocks = new();

    private void AddTables() {
        for (int pageIndex = 0; pageIndex < _source.Pages.Count; pageIndex++) AddPageTables(_source.Pages[pageIndex], pageIndex + 1);
        foreach (PublisherPage master in _source.MasterPages) AddPageTables(master, null);
    }

    private void AddPageTables(PublisherPage page, int? pageNumber) {
        for (int index = 0; index < page.Tables.Count; index++) {
            _token.ThrowIfCancellationRequested();
            PublisherTable native = page.Tables[index];
            // A recovered grid with unresolved native text boundaries remains in
            // PublisherPage.Tables; Reader does not turn unknown cell text into empty data.
            if (!native.HasTextMapping || !native.StoryId.HasValue) continue;
            Item();
            ReaderLocation location = Location(pageNumber.HasValue ? "publisher-table" : "publisher-master-table",
                "publisher-object-" + native.Id.ToString(CultureInfo.InvariantCulture), index);
            location.Page = pageNumber; location.TableIndex = index;
            int rowCount = Math.Min(native.RowCount, _settings.MaxTableRows);
            var columns = new string[native.ColumnCount];
            for (int column = 0; column < columns.Length; column++) {
                Item(); columns[column] = "Column " + (column + 1).ToString(CultureInfo.InvariantCulture);
                Characters(columns[column].Length);
            }
            var rows = new string[rowCount][];
            for (int row = 0; row < rows.Length; row++) {
                Item(); rows[row] = new string[native.ColumnCount];
                for (int column = 0; column < columns.Length; column++) { Item(); rows[row][column] = string.Empty; }
            }
            bool merged = false;
            foreach (PublisherTableCell cell in native.Cells) {
                Item();
                merged |= cell.RowSpan > 1 || cell.ColumnSpan > 1;
                if (cell.RowIndex >= rowCount) continue;
                Characters(cell.Text.Length);
                rows[cell.RowIndex][cell.ColumnIndex] = cell.Text;
            }
            var table = new ReaderTable {
                Title = "Publisher table " + native.Id.ToString(CultureInfo.InvariantCulture), Kind = "publisher-table",
                Location = location, Columns = columns, Rows = rows,
                ColumnProfiles = ReaderTableProfiler.CreateProfiles(columns, rows),
                TotalRowCount = native.RowCount, Truncated = rowCount < native.RowCount
            };
            _tables.Add(table);
            if (!_storyTables.TryGetValue(native.StoryId.Value, out var storyTables))
                _storyTables.Add(native.StoryId.Value, storyTables = new());
            storyTables.Add(table);
            if (pageNumber.HasValue) {
                if (!_pageTables.TryGetValue(page.Id, out var pageTables)) _pageTables.Add(page.Id, pageTables = new());
                pageTables.Add(table);
            }
            if (merged) TableDiagnostic("PUB_READER_TABLE_SPANS_FLATTENED",
                "Reader rows retain each merged cell's text at its first row and column. Covered positions are empty; the native spans remain in PublisherPage.Tables.",
                OfficeConversionLossKind.Approximation, location);
            if (table.Truncated) TableDiagnostic("PUB_READER_TABLE_ROWS_TRUNCATED",
                "Reader table rows were truncated by MaxTableRows. Complete source text remains in the table story block and chunks.",
                OfficeConversionLossKind.Omission, location);
        }
    }

    private void TableDiagnostic(string code, string message, OfficeConversionLossKind loss, ReaderLocation location) =>
        _diagnostics.Add(new OfficeDocumentDiagnostic {
            Code = code, Message = message, Source = "OfficeIMO.Reader.Publisher",
            Severity = OfficeDocumentDiagnosticSeverity.Warning, Category = OfficeDocumentDiagnosticCategory.Content,
            Location = location, Attributes = new Dictionary<string, string> { ["lossKind"] = loss.ToString() }
        });

    private void Characters(long count) {
        _token.ThrowIfCancellationRequested();
        if (count > _options.ReadOptions!.Limits.MaxTextCharacters - _characters)
            throw new InvalidDataException("Publisher Reader projection text limit exceeded.");
        _characters += count;
    }
}
