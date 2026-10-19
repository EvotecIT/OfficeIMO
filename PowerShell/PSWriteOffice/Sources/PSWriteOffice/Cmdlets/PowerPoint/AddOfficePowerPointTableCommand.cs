using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Management.Automation;
using OfficeIMO.PowerPoint;
using PSWriteOffice.Services;
using PSWriteOffice.Services.PowerPoint;
using PSWriteOffice.Services.Table;

namespace PSWriteOffice.Cmdlets.PowerPoint;

/// <summary>Adds a table to a PowerPoint slide.</summary>
/// <para>Builds a table from data rows or creates a blank grid with a fixed size.</para>
/// <example>
///   <summary>Create a table from objects.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$rows = @([pscustomobject]@{ Item='Alpha'; Qty=2 }, [pscustomobject]@{ Item='Beta'; Qty=4 })
///   Add-OfficePowerPointTable -Slide $slide -InputObject $rows -X 60 -Y 140 -Width 420 -Height 200</code>
///   <para>Creates a table with headers and two data rows.</para>
/// </example>
[Cmdlet(VerbsCommon.Add, "OfficePowerPointTable", DefaultParameterSetName = ParameterSetInputObject)]
[Alias("PptTable")]
public sealed class AddOfficePowerPointTableCommand : PSWriteOffice.Cmdlets.OfficeMutationCmdlet {
    private const string ParameterSetInputObject = "InputObject";
    private const string ParameterSetSize = "Size";

    /// <summary>Target slide that will receive the table (optional inside DSL).</summary>
    [Parameter(ValueFromPipeline = true, Position = 0)]
    public PowerPointSlide? Slide { get; set; }

    /// <summary>Source objects to convert into table rows.</summary>
    [Parameter(Mandatory = true, Position = 1, ParameterSetName = ParameterSetInputObject)]
    [Alias("Data")]
    public object? InputObject { get; set; }

    /// <summary>Optional header order to apply to the table.</summary>
    [Parameter(ParameterSetName = ParameterSetInputObject)]
    [Alias("Headers")]
    public string[]? Header { get; set; }

    /// <summary>Skip writing header row.</summary>
    [Parameter(ParameterSetName = ParameterSetInputObject)]
    public SwitchParameter NoHeader { get; set; }

    /// <summary>Projection to apply before writing the table.</summary>
    [Parameter(ParameterSetName = ParameterSetInputObject)]
    public OfficeTableView View { get; set; } = OfficeTableView.Normal;

    /// <summary>Text used between items when a cell contains a collection.</summary>
    [Parameter(ParameterSetName = ParameterSetInputObject)]
    [AllowEmptyString]
    public string CollectionSeparator { get; set; } = ", ";

    /// <summary>Text used between entries when a cell contains a dictionary.</summary>
    [Parameter(ParameterSetName = ParameterSetInputObject)]
    [AllowEmptyString]
    public string DictionaryEntrySeparator { get; set; } = "; ";

    /// <summary>Text used between a dictionary key and value.</summary>
    [Parameter(ParameterSetName = ParameterSetInputObject)]
    [AllowEmptyString]
    public string DictionaryKeyValueSeparator { get; set; } = ": ";

    /// <summary>Maximum number of items allowed in one nested collection or dictionary cell. Defaults to 1,048,575; increase explicitly for trusted larger values.</summary>
    [Parameter(ParameterSetName = ParameterSetInputObject)]
    [ValidateRange(1, int.MaxValue)]
    public int MaxCollectionItems { get; set; } = PowerShellObjectNormalizerOptions.DefaultMaxCollectionItems;

    /// <summary>Maximum nesting depth allowed while normalizing one cell value. Defaults to 64; increase explicitly for trusted deeper values.</summary>
    [Parameter(ParameterSetName = ParameterSetInputObject)]
    [ValidateRange(1, int.MaxValue)]
    public int MaxNestingDepth { get; set; } = PowerShellObjectNormalizerOptions.DefaultMaxNestingDepth;

    /// <summary>Row count for an empty table.</summary>
    [Parameter(Mandatory = true, ParameterSetName = ParameterSetSize)]
    public int Rows { get; set; }

    /// <summary>Column count for an empty table.</summary>
    [Parameter(Mandatory = true, ParameterSetName = ParameterSetSize)]
    public int Columns { get; set; }

    /// <summary>Left offset (in points) from the slide origin.</summary>
    [Parameter]
    public double X { get; set; } = 0;

    /// <summary>Top offset (in points) from the slide origin.</summary>
    [Parameter]
    public double Y { get; set; } = 0;

    /// <summary>Table width in points.</summary>
    [Parameter]
    public double Width { get; set; } = 400;

    /// <summary>Table height in points.</summary>
    [Parameter]
    public double Height { get; set; } = 240;

    /// <summary>Optional table style ID (GUID string).</summary>
    [Parameter]
    public string? StyleId { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        PowerPointSlide? slide = null;
        try {
            ValidateDimensions();
            slide = Slide ?? PowerPointDslContext.Require(this).RequireSlide();

            PowerPointTable table = ParameterSetName == ParameterSetSize
                ? CreateSizedTable(slide)
                : CreateDataTable(slide);

            if (!string.IsNullOrWhiteSpace(StyleId)) {
                table.StyleId = StyleId;
            }

            WritePassThru(table);
        } catch (Exception ex) {
            WriteError(new ErrorRecord(ex, "PowerPointAddTableFailed", ErrorCategory.InvalidOperation, slide ?? Slide));
        }
    }

    private PowerPointTable CreateSizedTable(PowerPointSlide slide) {
        if (Rows <= 0) {
            throw new ArgumentOutOfRangeException(nameof(Rows), "Rows must be greater than 0.");
        }

        if (Columns <= 0) {
            throw new ArgumentOutOfRangeException(nameof(Columns), "Columns must be greater than 0.");
        }

        return slide.AddTablePoints(Rows, Columns, X, Y, Width, Height);
    }

    private PowerPointTable CreateDataTable(PowerPointSlide slide) {
        var items = new List<object?>();
        TableInputCollector.AddInput(items, InputObject);
        var inputRows = TableInputCollector.RequireRows(items, nameof(InputObject));
        var projectedRows = TableViewProjection.Project(inputRows, View);
        var normalizerOptions = PowerShellObjectNormalizerOptions.ForTable(
            CollectionSeparator,
            DictionaryEntrySeparator,
            DictionaryKeyValueSeparator,
            MaxCollectionItems,
            MaxNestingDepth);
        var explicitHeaders = ResolveExplicitHeaders();
        if (OfficeTableSpecParser.TryCreate(
                projectedRows,
                propertyNames: explicitHeaders,
                header: NoHeader.IsPresent ? Array.Empty<string>() : explicitHeaders,
                out var tableSpec,
                normalizerOptions)) {
            return CreateStructuredTable(slide, tableSpec);
        }

        var normalized = PowerShellObjectNormalizer.NormalizeItems(projectedRows, normalizerOptions);
        var rows = NormalizeRows(normalized);
        var headers = ResolveHeaders(rows);

        if (headers.Count == 0) {
            throw new InvalidOperationException("Unable to infer columns from the supplied data.");
        }

        var columns = headers
            .Select(header => PowerPointTableColumn<Dictionary<string, object?>>.Create(
                header,
                row => row.TryGetValue(header, out var value) ? value : null))
            .ToList();

        return slide.AddTablePoints(rows, columns, includeHeaders: !NoHeader.IsPresent, X, Y, Width, Height);
    }

    private PowerPointTable CreateStructuredTable(PowerPointSlide slide, OfficeTableSpec spec) {
        foreach (var placement in spec.Placements) {
            PowerPointTableCellSpecService.Validate(placement.Cell);
        }

        var table = slide.AddTablePoints(spec.RowCount, spec.ColumnCount, X, Y, Width, Height);
        foreach (var placement in spec.Placements) {
            var cell = table.GetCell(placement.RowIndex, placement.ColumnIndex);
            PowerPointTableCellSpecService.Apply(cell, placement.Cell);
            if (placement.Cell.HasSpan) {
                cell.Merge = (placement.Cell.RowSpan, placement.Cell.ColumnSpan);
            }
        }

        return table;
    }

    private void ValidateDimensions() {
        if (Width <= 0) {
            throw new ArgumentOutOfRangeException(nameof(Width), "Width must be greater than 0.");
        }

        if (Height <= 0) {
            throw new ArgumentOutOfRangeException(nameof(Height), "Height must be greater than 0.");
        }
    }

    private List<Dictionary<string, object?>> NormalizeRows(IReadOnlyList<object?> items) {
        var rows = new List<Dictionary<string, object?>>(items.Count);
        foreach (var item in items) {
            if (item == null) {
                rows.Add(new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase));
                continue;
            }

            if (item is IDictionary dict) {
                var row = new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase);
                foreach (DictionaryEntry entry in dict) {
                    var key = Convert.ToString(entry.Key, CultureInfo.InvariantCulture);
                    if (string.IsNullOrWhiteSpace(key)) {
                        continue;
                    }
                    row[key] = entry.Value;
                }

                rows.Add(row);
                continue;
            }

            rows.Add(new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase) {
                ["Value"] = item
            });
        }

        return rows;
    }

    private List<string> ResolveHeaders(IReadOnlyList<Dictionary<string, object?>> rows) {
        var explicitHeaders = ResolveExplicitHeaders();
        if (explicitHeaders != null) {
            return explicitHeaders.ToList();
        }

        var headers = new List<string>();
        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var row in rows) {
            foreach (var key in row.Keys) {
                if (seen.Add(key)) {
                    headers.Add(key);
                }
            }
        }

        return headers;
    }

    private string[]? ResolveExplicitHeaders() {
        if (Header == null || Header.Length == 0) {
            return null;
        }

        var explicitHeaders = Header
            .Where(h => !string.IsNullOrWhiteSpace(h))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToArray();

        if (explicitHeaders.Length == 0) {
            throw new PSArgumentException("Header cannot be empty.", nameof(Header));
        }

        return explicitHeaders;
    }
}