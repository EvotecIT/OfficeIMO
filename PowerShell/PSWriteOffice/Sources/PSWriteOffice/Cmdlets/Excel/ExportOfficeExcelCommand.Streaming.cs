using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Management.Automation;
using System.Threading;
using OfficeIMO;
using OfficeIMO.Excel;
using PSWriteOffice.Services;

namespace PSWriteOffice.Cmdlets.Excel;

public sealed partial class ExportOfficeExcelCommand {
    private readonly CancellationTokenSource _streamCancellation = new();
    private ExcelWorkbookWriter? _streamWriter;
    private ExcelIncrementalTableWriter? _streamRows;
    private PowerShellObjectNormalizerOptions? _streamNormalizer;
    private ExcelColumnFormatPlan? _streamFormats;
    private HashSet<string>? _streamExcluded;
    private readonly HashSet<string> _streamHeaders = new(StringComparer.OrdinalIgnoreCase);
    private string? _streamPath;
    private bool _streamDeclined;
    private bool _streamHasInput;
    private bool _streamDataSet;
    private bool _streamDisposed;
    private bool StreamingCreate => !Append.IsPresent && !ReplaceSheet.IsPresent;

    /// <inheritdoc />
    protected override void BeginProcessing() {
        if (!StreamingCreate) return;
        _streamPath = SessionState.Path.GetUnresolvedProviderPathFromPSPath(Path);
        if (NoClobber && File.Exists(_streamPath)) throw new IOException($"File '{_streamPath}' already exists.");
        _ = StartRow;
        _ = StartColumn;
        _streamFormats = BuildColumnFormatPlan();
        _streamNormalizer = CreateNormalizerOptions();
        _streamExcluded = ExcludeProperty == null ? null : new HashSet<string>(
            ExcludeProperty.Where(name => !string.IsNullOrWhiteSpace(name)).Select(name => name.Trim()),
            StringComparer.OrdinalIgnoreCase);
        _streamDeclined = !ShouldProcess(_streamPath, File.Exists(_streamPath) ? "Overwrite Excel workbook" : "Write new Excel workbook");
    }

    private void EnsureStreamingWriter(string? inferredTableName = null, bool dataSet = false) {
        if (_streamWriter != null) return;
        if (!dataSet && Sheet != null) {
            if (Sheet.Sheet != null || Sheet.Position is > 1)
                throw new PSArgumentException("A new workbook accepts a worksheet name or position 1.", nameof(Sheet));
            _resolvedSheetName = Sheet.Position == 1 ? "Sheet1" : Sheet.Name;
        }
        var properties = DocumentProperty?.Cast<DictionaryEntry>().ToDictionary<DictionaryEntry, string, object?>(
            entry => Convert.ToString(entry.Key, CultureInfo.InvariantCulture)!,
            entry => entry.Value is PSObject wrapped ? wrapped.BaseObject : entry.Value,
            StringComparer.OrdinalIgnoreCase);
        var options = new ExcelWriterOptions {
            Address = Address ?? "A1", Title = Title, CreateTable = !NoTable, IncludeHeaders = !NoHeader,
            TableName = dataSet ? null : TableName ?? inferredTableName,
            TableStyle = NoTableStyle ? ExcelTableStyle.None : TableStyle,
            IncludeAutoFilter = !NoAutoFilter, AutoFit = !NoAutoFit, BoldHeaders = !NoHeader,
            FreezeHeader = !NoFreeze && !NoHeader, FreezeFirstColumn = FreezeFirstColumn,
            ShowFirstColumn = ShowFirstColumn, ShowLastColumn = ShowLastColumn,
            ShowColumnStripes = ShowColumnStripes, ShowRowStripes = !NoRowStripes,
            ColumnFormat = _streamFormats, IgnoreMissingColumnFormat = IgnoreMissingColumnFormat || dataSet,
            DocumentProperties = properties, SaveOptions = SaveOptions,
            DateSystem = DateSystem ?? ExcelDateSystem.NineteenHundred, NoClobber = NoClobber,
            UseSharedStrings = false, UseCellValueNumberFormats = true
        };
        _streamWriter = ExcelDocument.CreateWriter(_streamPath!, options, Theme?.Resolve(OfficeTheme.FromDocument));
    }

    private void StreamInput(object? value) {
        _streamCancellation.Token.ThrowIfCancellationRequested();
        if (value is PSObject wrapped && wrapped.BaseObject is not PSCustomObject) value = wrapped.BaseObject;
        if (value == null) return;
        if (_streamDataSet) throw new PSArgumentException("A DataSet must be the only input to this export.", nameof(InputObject));
        if (value is DataSet dataSet) {
            if (_streamHasInput || dataSet.Tables.Count == 0)
                throw new PSArgumentException("Supply one nonempty DataSet as the only export input.", nameof(InputObject));
            _streamDataSet = true;
            _streamHasInput = true;
            EnsureStreamingWriter(dataSet: true);
            string[] names = BuildDataSetWorksheetNames(dataSet, Array.Empty<string>());
            for (int index = 0; index < names.Length; index++) {
                _streamRows = _streamWriter!.BeginTable(names[index], _streamCancellation.Token);
                StreamDataTable(dataSet.Tables[index]);
                CompleteStreamingTable();
            }
            return;
        }
        if (value is DataTable table) {
            EnsureStreamingRows(table.TableName);
            _streamHasInput = true;
            StreamDataTable(table);
        } else if (value is DataView view) {
            EnsureStreamingRows(view.Table?.TableName);
            _streamHasInput = true;
            DeclareStreamingColumns(view.Table?.Columns.Cast<DataColumn>().Select(column => column.ColumnName)
                ?? Array.Empty<string>());
            foreach (DataRowView row in view) StreamProjectedRow(row);
        } else if (value is IDataReader reader) {
            EnsureStreamingRows("ReaderData");
            _streamHasInput = true;
            string[] headers = ExcelSheet.BuildReaderHeaders(reader);
            DeclareStreamingColumns(headers);
            while (reader.Read()) {
                _streamCancellation.Token.ThrowIfCancellationRequested();
                var values = new object?[headers.Length];
                for (int column = 0; column < values.Length; column++)
                    values[column] = reader.IsDBNull(column) ? null :
                        PowerShellObjectNormalizer.NormalizeCellValueForTable(reader.GetValue(column), _streamNormalizer);
                WriteStreamingRow(headers, values);
            }
        } else if (value is IEnumerable sequence && value is not string &&
                   !PowerShellDictionaryAdapter.IsDictionaryLike(value)) {
            foreach (object? item in sequence) StreamInput(item);
        } else {
            EnsureStreamingRows();
            _streamHasInput = true;
            StreamProjectedRow(value);
        }
    }

    private void EnsureStreamingRows(string? tableName = null) {
        EnsureStreamingWriter(string.IsNullOrWhiteSpace(tableName) ? null : tableName);
        _streamRows ??= _streamWriter!.BeginTable(WorksheetName, _streamCancellation.Token);
    }

    private void StreamDataTable(DataTable table) {
        DeclareStreamingColumns(table.Columns.Cast<DataColumn>().Select(column => column.ColumnName));
        foreach (DataRow row in table.Rows) StreamProjectedRow(row);
    }

    private void DeclareStreamingColumns(IEnumerable<string> headers) {
        string[] columns = headers.Where(header => _streamExcluded?.Contains(header) != true).ToArray();
        _streamRows!.DeclareColumns(columns);
        _streamHeaders.UnionWith(columns);
    }

    private void StreamProjectedRow(object value) {
        _streamCancellation.Token.ThrowIfCancellationRequested();
        if (PowerShellObjectNormalizer.TryProjectItem(value, null, out string[] columns, out object?[] values, _streamNormalizer))
            WriteStreamingRow(columns, values);
    }

    private void WriteStreamingRow(string[] columns, object?[] values) {
        if (_streamExcluded is { Count: > 0 }) {
            int[] selected = Enumerable.Range(0, columns.Length).Where(index => !_streamExcluded.Contains(columns[index])).ToArray();
            values = selected.Select(index => values[index]).ToArray();
            columns = selected.Select(index => columns[index]).ToArray();
        }
        _streamHeaders.UnionWith(columns);
        _streamRows!.WriteRow(columns, values);
    }

    private void CompleteStreamingTable() {
        ExcelDataSetImportResult result = _streamRows!.Complete();
        _streamRows = null;
        WriteVerbose($"Exported data to {result.SheetName}!{result.Range} through the native streaming writer.");
    }

    private void CompleteStreamingExport() {
        if (_streamDeclined) return;
        try {
            _streamCancellation.Token.ThrowIfCancellationRequested();
            if (!_streamHasInput) throw new PSArgumentException("Provide at least one object to export.", nameof(InputObject));
            if (_streamHeaders.Count == 0) throw new PSArgumentException("Unable to infer columns from the supplied data.", nameof(InputObject));
            if (_streamDataSet) ThrowForDataSetColumnFormatsMissingEverywhere(_streamFormats, _streamHeaders);
            if (_streamRows != null) CompleteStreamingTable();
            _streamWriter!.Dispose();
            _streamWriter = null;
            if (Show) FileOpenService.Open(_streamPath!);
            WritePassThru(_streamPath!);
        } catch { AbortStreamingExport(); throw; }
    }

    private void AbortStreamingExport() {
        ExcelWorkbookWriter? writer = _streamWriter;
        _streamWriter = null;
        _streamRows = null;
        if (writer == null) return;
        writer.Abort();
        writer.Dispose();
    }

    /// <inheritdoc />
    protected override void StopProcessing() => _streamCancellation.Cancel();

    /// <summary>Discards an interrupted export and releases its native row spool.</summary>
    public void Dispose() {
        if (_streamDisposed) return;
        _streamDisposed = true;
        try { AbortStreamingExport(); }
        finally { _streamCancellation.Dispose(); }
    }
}
