using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Management.Automation;
using OfficeIMO;
using OfficeIMO.PowerPoint;
using PSWriteOffice.Foundation;
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
[Alias("PowerPointTable")]
[OutputType(typeof(PowerPointTable))]
public sealed class AddOfficePowerPointTableCommand : PSWriteOffice.Cmdlets.OfficeMutationCmdlet {
    private const string ParameterSetInputObject = "InputObject";
    private const string ParameterSetSize = "Size";

    /// <summary>Target slide that will receive the table (optional inside DSL).</summary>
    [Parameter, PowerPointSlideTransformation]
    public PowerPointSlideSelector? Slide { get; set; }

    /// <summary>Source objects to convert into table rows.</summary>
    [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true, ParameterSetName = ParameterSetInputObject)]
    public object?[]? InputObject { get; set; }

    /// <summary>Optional header order to apply to the table.</summary>
    [Parameter(ParameterSetName = ParameterSetInputObject)]
    public string[]? Property { get; set; }

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

    /// <summary>Maximum data rows per physical slide; omission uses the shared layout capacity.</summary>
    [Parameter(ParameterSetName = ParameterSetInputObject), ValidateRange(1, int.MaxValue)]
    public int? MaxRows { get; set; }

    /// <summary>Keep the table together instead of generating continuation slides.</summary>
    [Parameter(ParameterSetName = ParameterSetInputObject)]
    public SwitchParameter NoSplit { get; set; }

    /// <summary>Row count for an empty table.</summary>
    [Parameter(Mandatory = true, ParameterSetName = ParameterSetSize)]
    public int Rows { get; set; }

    /// <summary>Column count for an empty table.</summary>
    [Parameter(Mandatory = true, ParameterSetName = ParameterSetSize)]
    public int Columns { get; set; }

    /// <summary>Left offset from the slide origin; numbers mean points.</summary>
    [Parameter, OfficeLengthTransformation]
    public OfficeLength X { get; set; }

    /// <summary>Top offset from the slide origin; numbers mean points.</summary>
    [Parameter, OfficeLengthTransformation]
    public OfficeLength Y { get; set; }

    /// <summary>Table width; numbers mean points.</summary>
    [Parameter, OfficeLengthTransformation]
    public OfficeLength Width { get; set; } = OfficeLength.Points(400);

    /// <summary>Table height; numbers mean points.</summary>
    [Parameter, OfficeLengthTransformation]
    public OfficeLength Height { get; set; } = OfficeLength.Points(240);

    /// <summary>Native table style name or GUID; omission retains the native theme style.</summary>
    [Parameter]
    public string? Style { get; set; }

    /// <summary>Culture for formatting data; lengths always parse invariantly.</summary>
    [Parameter(ParameterSetName = ParameterSetInputObject), ValidateNotNull]
    public CultureInfo Culture { get; set; } = CultureInfo.CurrentCulture;

    /// <summary>One-based start column in the native twelve-column grid; explicit X overrides it.</summary>
    [Parameter, ValidateRange(1, 12)]
    public int? Column { get; set; }

    /// <summary>Number of grid columns; defaults to twelve, and explicit Width overrides it.</summary>
    [Parameter, ValidateRange(1, 12)]
    public int? Span { get; set; }

    private readonly List<object?> _input = new();
    private OfficeLength _left;
    private OfficeLength _width;

    /// <inheritdoc />
    protected override void ProcessRecord() {
        if (ParameterSetName == ParameterSetInputObject) TableInputCollector.AddInput(_input, InputObject);
    }

    /// <inheritdoc />
    protected override void EndProcessing() {
        if (ParameterSetName == ParameterSetInputObject) InputObject = _input.ToArray();
        PowerPointSlide? slide = null;
        PowerPointTable? table = null;
        try {
            ValidateDimensions();
            slide = PowerPointManualPlacement.Resolve(Slide).Slide;
            var horizontal = PowerPointManualPlacement.Horizontal(slide,
                MyInvocation.BoundParameters.ContainsKey(nameof(X)) ? X : null,
                MyInvocation.BoundParameters.ContainsKey(nameof(Width)) ? Width : null, Column, Span);
            _left = horizontal.X ?? X;
            _width = horizontal.Width ?? Width;
            bool explicitBounds = PowerPointManualPlacement.HasExplicitBounds(MyInvocation.BoundParameters.Keys) || Column.HasValue || Span.HasValue;
            if (Slide == null && !explicitBounds && PowerPointContentScope.Current != null && ParameterSetName == ParameterSetInputObject) {
                AddLayoutTable();
                return;
            }
            table = ParameterSetName == ParameterSetSize
                ? CreateSizedTable(slide)
                : CreateDataTable(slide);

            ApplyStyle(table);

            var scope = PowerPointContentScope.Current;
            if (scope != null && ReferenceEquals(scope.NativeSlide, slide)) scope.TrackFixedFigure(table);
            WritePassThru(table);
        } catch (Exception ex) {
            table?.Remove();
            if (ex is PipelineStoppedException or ActionPreferenceStopException) throw;
            if (PowerPointContentScope.Current != null) ThrowTerminatingError(OfficeError.Create(ex, "invalid-argument", InputObject, ErrorCategory.InvalidArgument));
            WriteError(new ErrorRecord(ex, "PowerPointAddTableFailed", ErrorCategory.InvalidOperation, slide ?? (object?)Slide));
        }
    }

    private void AddLayoutTable() {
        ValidateDimensions();
        var scope = PowerPointContentScope.Current!;
        PowerPointTable table = CreateDataTable(scope.NativeSlide);
        try {
            ApplyStyle(table);
            var native = new PowerPointLayoutNativeTable(table, !NoHeader);
            scope.AddTable(new PowerPointLayoutTable(native.Headers, native.Rows) {
                IncludeHeaders = !NoHeader, MaximumRows = MaxRows, NoSplit = NoSplit, NativeBody = native
            });
        } catch {
            table.Remove();
            throw;
        }
        WritePassThru(table);
    }

    private void ApplyStyle(PowerPointTable table) {
        if (string.IsNullOrWhiteSpace(Style)) return;
        if (Guid.TryParse(Style, out Guid id)) table.StyleId = id.ToString("B").ToUpperInvariant();
        else table.ApplyStyleByName(Style!, ignoreCase: true);
    }

    private PowerPointTable CreateSizedTable(PowerPointSlide slide) {
        if (Rows <= 0) {
            throw new ArgumentOutOfRangeException(nameof(Rows), "Rows must be greater than 0.");
        }

        if (Columns <= 0) {
            throw new ArgumentOutOfRangeException(nameof(Columns), "Columns must be greater than 0.");
        }

        return slide.AddTable(Rows, Columns, _left, Y, _width, Height);
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
        normalizerOptions.Culture = Culture;
        normalizerOptions.FormatScalarValuesAsText = true;
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

        return slide.AddTable(rows, columns, includeHeaders: !NoHeader.IsPresent, _left, Y, _width, Height);
    }

    private PowerPointTable CreateStructuredTable(PowerPointSlide slide, OfficeTableSpec spec) {
        foreach (var placement in spec.Placements) {
            PowerPointTableCellSpecService.Validate(placement.Cell);
        }

        var table = slide.AddTable(spec.RowCount, spec.ColumnCount, _left, Y, _width, Height);
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
        if (Width <= OfficeLength.Zero) {
            throw new ArgumentOutOfRangeException(nameof(Width), "Width must be greater than 0.");
        }

        if (Height <= OfficeLength.Zero) {
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
        if (Property == null || Property.Length == 0) {
            return null;
        }

        var explicitHeaders = Property
            .Where(h => !string.IsNullOrWhiteSpace(h))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToArray();

        if (explicitHeaders.Length == 0) {
            throw new PSArgumentException("Property cannot be empty.", nameof(Property));
        }

        return explicitHeaders;
    }
}
