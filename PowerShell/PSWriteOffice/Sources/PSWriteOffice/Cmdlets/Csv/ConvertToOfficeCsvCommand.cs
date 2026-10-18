using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;
using System.Management.Automation;
using OfficeIMO.CSV;

namespace PSWriteOffice.Cmdlets.Csv;

/// <summary>Converts objects or a CSV document into CSV text.</summary>
/// <para>Use <c>Export-OfficeCsv</c> when the destination is a file.</para>
/// <example>
///   <summary>Convert objects to CSV text.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$csv = $data | ConvertTo-OfficeCsv</code>
///   <para>Generates CSV text from the input objects.</para>
/// </example>
/// <example>
///   <summary>Export with a stable schema order.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$rows = @(
///     [ordered]@{ Id = 1; Name = 'Alpha'; Total = 10.5 },
///     [ordered]@{ Id = 2; Name = 'Beta'; Total = 7.25 }
///   )
///   $csv = $rows | ConvertTo-OfficeCsv -Delimiter ';'</code>
///   <para>Uses ordered dictionaries to enforce column order and a custom delimiter.</para>
/// </example>
/// <example>
///   <summary>Write CSV without headers.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$csv = $data | ConvertTo-OfficeCsv -NoHeader</code>
///   <para>Writes rows only when a downstream system expects headerless CSV.</para>
/// </example>
[Cmdlet(VerbsData.ConvertTo, "OfficeCsv", DefaultParameterSetName = ParameterSetInputObjectDelimiter)]
[OutputType(typeof(string))]
public sealed class ConvertToOfficeCsvCommand : PSCmdlet
{
    private const string ParameterSetInputObjectDelimiter = "InputObjectDelimiter";
    private const string ParameterSetInputObjectCulture = "InputObjectCulture";
    private const string ParameterSetDocumentDelimiter = "DocumentDelimiter";
    private const string ParameterSetDocumentCulture = "DocumentCulture";
    private readonly CsvPowerShellObjectProjector _objectProjector = new();
    private CsvPowerShellLineWriter? _lineWriter;
    private CsvRowWriter? _csvWriter;

    /// <summary>CSV document to serialize.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = ParameterSetDocumentDelimiter)]
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = ParameterSetDocumentCulture)]
    public CsvDocument Document { get; set; } = null!;

    /// <summary>Objects to convert into CSV rows.</summary>
    [Parameter(ValueFromPipeline = true, ParameterSetName = ParameterSetInputObjectDelimiter)]
    [Parameter(ValueFromPipeline = true, ParameterSetName = ParameterSetInputObjectCulture)]
    public object? InputObject { get; set; }

    /// <summary>Field delimiter character.</summary>
    [Parameter(ParameterSetName = ParameterSetInputObjectDelimiter)]
    [Parameter(ParameterSetName = ParameterSetDocumentDelimiter)]
    public char Delimiter { get; set; } = ',';

    /// <summary>Field delimiter text for multi-character delimiters such as || or ::.</summary>
    [Parameter(ParameterSetName = ParameterSetInputObjectDelimiter)]
    [Parameter(ParameterSetName = ParameterSetDocumentDelimiter)]
    public string? DelimiterText { get; set; }

    /// <summary>Use the list separator from the selected or current culture as the delimiter.</summary>
    [Parameter(Mandatory = true, ParameterSetName = ParameterSetInputObjectCulture)]
    [Parameter(Mandatory = true, ParameterSetName = ParameterSetDocumentCulture)]
    public SwitchParameter UseCulture { get; set; }

    /// <summary>Omit the header row from the output.</summary>
    [Parameter]
    public SwitchParameter NoHeader { get; set; }

    /// <summary>Override the newline sequence.</summary>
    [Parameter]
    public string? NewLine { get; set; }

    /// <summary>Culture used for value formatting.</summary>
    [Parameter]
    public CultureInfo? Culture { get; set; }

    /// <summary>Controls how formula-like values are written.</summary>
    [Parameter]
    public CsvFormulaInjectionPolicy FormulaInjectionPolicy { get; set; } = CsvFormulaInjectionPolicy.Preserve;

    /// <summary>Controls when CSV fields are quoted. Defaults to quoting only fields that need it.</summary>
    [Parameter]
    public CsvQuoteMode UseQuotes { get; set; } = CsvQuoteMode.AsNeeded;

    /// <summary>Field names that should always be quoted when <see cref="UseQuotes"/> is AsNeeded.</summary>
    [Parameter]
    public string[]? QuoteFields { get; set; }

    /// <summary>Token written for null values.</summary>
    [Parameter]
    public string? NullValue { get; set; }

    /// <summary>Date/time format used for DateTime and DateTimeOffset values.</summary>
    [Parameter]
    public string? DateTimeFormat { get; set; }

    /// <summary>Convert date/time values to UTC before formatting.</summary>
    [Parameter]
    public SwitchParameter UseUtc { get; set; }

    /// <summary>Maximum number of items allowed in one nested collection or dictionary field. Defaults to 1,048,575; increase explicitly for trusted larger values.</summary>
    [Parameter(ParameterSetName = ParameterSetInputObjectDelimiter)]
    [Parameter(ParameterSetName = ParameterSetInputObjectCulture)]
    [ValidateRange(1, int.MaxValue)]
    public int MaxCollectionItems { get; set; } = PSWriteOffice.Services.PowerShellObjectNormalizerOptions.DefaultMaxCollectionItems;

    /// <summary>Maximum nesting depth allowed while normalizing one field value. Defaults to 64; increase explicitly for trusted deeper values.</summary>
    [Parameter(ParameterSetName = ParameterSetInputObjectDelimiter)]
    [Parameter(ParameterSetName = ParameterSetInputObjectCulture)]
    [ValidateRange(1, int.MaxValue)]
    public int MaxNestingDepth { get; set; } = PSWriteOffice.Services.PowerShellObjectNormalizerOptions.DefaultMaxNestingDepth;

    /// <inheritdoc />
    protected override void BeginProcessing()
    {
        ApplyCultureDelimiter();
    }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        if (IsDocumentParameterSet())
        {
            if (Document != null)
            {
                EmitCsv(Document);
            }

            return;
        }

        if (TryGetCsvDocument(InputObject, out var csvDocument))
        {
            EmitCsv(csvDocument);
            return;
        }

        if (TryGetDataTable(InputObject, out var dataTable))
        {
            EmitDataTable(dataTable);
            return;
        }

        _objectProjector.WriteObject(InputObject, EnsureObjectWriter());
    }

    /// <inheritdoc />
    protected override void EndProcessing()
    {
        DisposeObjectWriter();
    }

    private void ApplyCultureDelimiter()
    {
        if (!UseCulture.IsPresent)
        {
            return;
        }

        var separator = (Culture ?? CultureInfo.CurrentCulture).TextInfo.ListSeparator;
        if (string.IsNullOrEmpty(separator) || separator.Length != 1)
        {
            throw new PSArgumentException("The selected culture must provide a single-character list separator.");
        }

        Delimiter = separator[0];
    }

    private void EmitCsv(CsvDocument document)
    {
        var options = CreateSaveOptions();
        using var writer = new CsvPowerShellLineWriter(this, GetDelimiterText(options), options.QuoteMode);
        writer.Write(document.ToString(options));
    }

    private void EmitDataTable(DataTable table)
    {
        if (_csvWriter != null)
        {
            var activeColumns = _objectProjector.CurrentColumns ?? GetDataTableColumnNames(table);
            WriteDataTableRows(table, _csvWriter, activeColumns);
            return;
        }

        var tableColumns = GetDataTableColumnNames(table);
        _objectProjector.UseColumns(tableColumns, validateColumns: false);
        using var reader = table.CreateDataReader();
        EnsureObjectWriter().WriteDataReader(_objectProjector.NormalizeDataReader(reader));
    }

    private static string[] GetDataTableColumnNames(DataTable table)
    {
        var columns = new string[table.Columns.Count];
        for (var i = 0; i < columns.Length; i++)
        {
            columns[i] = table.Columns[i].ColumnName;
        }

        return columns;
    }

    private void WriteDataTableRows(DataTable table, CsvRowWriter writer, IReadOnlyList<string> columns)
    {
        foreach (DataRow row in table.Rows)
        {
            writer.WriteRow(
                columns,
                columns.Count,
                (Row: row, Columns: columns, Projector: _objectProjector),
                static (state, index) => state.Projector.NormalizeCellValue(
                    TryGetDataTableValue(state.Row, state.Columns[index])));
        }
    }

    private static object? TryGetDataTableValue(DataRow row, string column)
    {
        if (!TryGetDataColumn(row.Table, column, out var dataColumn))
        {
            return null;
        }

        var value = row[dataColumn];
        return value == DBNull.Value ? null : value;
    }

    private static bool TryGetDataColumn(DataTable table, string column, out DataColumn dataColumn)
    {
        if (table.Columns.Contains(column))
        {
            dataColumn = table.Columns[column]!;
            return true;
        }

        foreach (DataColumn candidate in table.Columns)
        {
            if (string.Equals(candidate.ColumnName, column, StringComparison.OrdinalIgnoreCase))
            {
                dataColumn = candidate;
                return true;
            }
        }

        dataColumn = null!;
        return false;
    }

    private static bool TryGetCsvDocument(object? value, out CsvDocument document)
    {
        if (value is CsvDocument csvDocument)
        {
            document = csvDocument;
            return true;
        }

        if (value is PSObject { BaseObject: CsvDocument psObjectDocument })
        {
            document = psObjectDocument;
            return true;
        }

        document = null!;
        return false;
    }

    private static bool TryGetDataTable(object? value, out DataTable table)
    {
        if (value is DataTable dataTable)
        {
            table = dataTable;
            return true;
        }

        if (value is PSObject { BaseObject: DataTable psObjectTable })
        {
            table = psObjectTable;
            return true;
        }

        table = null!;
        return false;
    }

    private CsvRowWriter EnsureObjectWriter()
    {
        if (_csvWriter != null)
        {
            return _csvWriter;
        }

        var options = CreateSaveOptions();
        _objectProjector.UseCsvOptions(options, MaxCollectionItems, MaxNestingDepth);
        _lineWriter = new CsvPowerShellLineWriter(this, GetDelimiterText(options), options.QuoteMode);
        _csvWriter = new CsvRowWriter(_lineWriter, options);
        return _csvWriter;
    }

    private void DisposeObjectWriter()
    {
        _csvWriter?.Dispose();
        _csvWriter = null;
        _lineWriter = null;
        _objectProjector.Reset();
    }

    private CsvSaveOptions CreateSaveOptions()
    {
        var options = new CsvSaveOptions
        {
            Delimiter = Delimiter,
            DelimiterText = DelimiterText,
            IncludeHeader = !NoHeader.IsPresent,
            Culture = Culture ?? CultureInfo.InvariantCulture,
            FormulaInjectionPolicy = FormulaInjectionPolicy,
            QuoteMode = UseQuotes,
            QuoteFields = QuoteFields
        };

        CsvPowerShellOptionBuilder.ApplySaveOptions(
            options,
            NullValue,
            DateTimeFormat,
            UseUtc.IsPresent,
            CsvCompressionType.None,
            System.IO.Compression.CompressionLevel.Optimal);

        if (!string.IsNullOrEmpty(NewLine))
        {
            options.NewLine = NewLine!;
        }

        return options;
    }

    private static string GetDelimiterText(CsvSaveOptions options) =>
        string.IsNullOrEmpty(options.DelimiterText) ? options.Delimiter.ToString() : options.DelimiterText!;

    private bool IsDocumentParameterSet() =>
        string.Equals(ParameterSetName, ParameterSetDocumentDelimiter, StringComparison.OrdinalIgnoreCase) ||
        string.Equals(ParameterSetName, ParameterSetDocumentCulture, StringComparison.OrdinalIgnoreCase);
}
