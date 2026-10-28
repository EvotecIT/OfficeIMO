using System.Management.Automation;
using OfficeIMO.Excel;
using PSWriteOffice.Services.Excel;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Applies a friendly AutoFilter condition by header name.</summary>
/// <example>
///   <summary>Filter a worksheet by header.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>ExcelSheet 'Data' { Set-OfficeExcelAutoFilter -Range A1:D200 -Header Status -Value Open,Hold }</code>
///   <para>Ensures an AutoFilter range exists and filters the Status column.</para>
/// </example>
[Cmdlet(VerbsCommon.Set, "OfficeExcelAutoFilter", DefaultParameterSetName = ParameterSetContext)]
[Alias("ExcelAutoFilterSet")]
public sealed class SetOfficeExcelAutoFilterCommand : PSCmdlet
{
    private const string ParameterSetContext = "Context";
    private const string ParameterSetDocument = "Document";

    /// <summary>Workbook to operate on outside the DSL context.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = ParameterSetDocument)]
    public ExcelDocument Document { get; set; } = null!;

    /// <summary>Worksheet name when using <see cref="Document"/>.</summary>
    [Parameter(ParameterSetName = ParameterSetDocument)]
    public string? Sheet { get; set; }

    /// <summary>Worksheet index (0-based) when using <see cref="Document"/>.</summary>
    [Parameter(ParameterSetName = ParameterSetDocument)]
    public int? SheetIndex { get; set; }

    /// <summary>Optional A1 AutoFilter range to create or replace before applying the condition.</summary>
    [Parameter]
    public string? Range { get; set; }

    /// <summary>Header name to filter.</summary>
    [Parameter(Mandatory = true, Position = 0)]
    [Alias("HeaderName", "ColumnName")]
    public string Header { get; set; } = string.Empty;

    /// <summary>Allowed values for an equals filter.</summary>
    [Parameter]
    [Alias("Values")]
    public string[]? Value { get; set; }

    /// <summary>Text that the column value must contain.</summary>
    [Parameter]
    public string? Contains { get; set; }

    /// <summary>Text that the column value must not contain.</summary>
    [Parameter]
    public string? DoesNotContain { get; set; }

    /// <summary>Text that the column value must start with.</summary>
    [Parameter]
    public string? StartsWith { get; set; }

    /// <summary>Text that the column value must end with.</summary>
    [Parameter]
    public string? EndsWith { get; set; }

    /// <summary>Numeric greater-than-or-equal condition.</summary>
    [Parameter]
    public double? GreaterThanOrEqual { get; set; }

    /// <summary>Numeric less-than-or-equal condition.</summary>
    [Parameter]
    public double? LessThanOrEqual { get; set; }

    /// <summary>Numeric not-equal condition.</summary>
    [Parameter]
    public double? NotEqual { get; set; }

    /// <summary>Inclusive numeric range condition. Provide exactly two values: minimum, maximum.</summary>
    [Parameter]
    public double[]? Between { get; set; }

    /// <summary>Outside numeric range condition. Provide exactly two values: minimum, maximum.</summary>
    [Parameter]
    public double[]? NotBetween { get; set; }

    /// <summary>Emit the worksheet after applying the filter.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        if (string.IsNullOrWhiteSpace(Header))
        {
            throw new PSArgumentException("Header cannot be empty.");
        }

        ValidateSingleCondition();

        var sheet = ResolveSheet();
        if (!string.IsNullOrWhiteSpace(Range))
        {
            sheet.AutoFilterAdd(Range!);
        }

        ApplyFilter(sheet);

        if (PassThru.IsPresent)
        {
            WriteObject(sheet);
        }
    }

    private ExcelSheet ResolveSheet()
    {
        if (ParameterSetName == ParameterSetDocument)
        {
            if (Document == null)
            {
                throw new PSArgumentException("Provide an Excel document.");
            }

            return ExcelSheetResolver.Resolve(Document, Sheet, SheetIndex);
        }

        var context = ExcelDslContext.Require(this);
        return context.RequireSheet();
    }

    private void ApplyFilter(ExcelSheet sheet)
    {
        if (Value is { Length: > 0 })
        {
            sheet.AutoFilterByHeaderEquals(Header, Value);
        }
        else if (!string.IsNullOrWhiteSpace(Contains))
        {
            sheet.AutoFilterByHeaderContains(Header, Contains!);
        }
        else if (!string.IsNullOrWhiteSpace(DoesNotContain))
        {
            sheet.AutoFilterByHeaderDoesNotContain(Header, DoesNotContain!);
        }
        else if (!string.IsNullOrWhiteSpace(StartsWith))
        {
            sheet.AutoFilterByHeaderStartsWith(Header, StartsWith!);
        }
        else if (!string.IsNullOrWhiteSpace(EndsWith))
        {
            sheet.AutoFilterByHeaderEndsWith(Header, EndsWith!);
        }
        else if (GreaterThanOrEqual.HasValue)
        {
            sheet.AutoFilterByHeaderGreaterThanOrEqual(Header, GreaterThanOrEqual.Value);
        }
        else if (LessThanOrEqual.HasValue)
        {
            sheet.AutoFilterByHeaderLessThanOrEqual(Header, LessThanOrEqual.Value);
        }
        else if (NotEqual.HasValue)
        {
            sheet.AutoFilterByHeaderNotEqual(Header, NotEqual.Value);
        }
        else if (Between is { Length: 2 })
        {
            sheet.AutoFilterByHeaderBetween(Header, Between[0], Between[1]);
        }
        else if (NotBetween is { Length: 2 })
        {
            sheet.AutoFilterByHeaderNotBetween(Header, NotBetween[0], NotBetween[1]);
        }
    }

    private void ValidateSingleCondition()
    {
        var count = 0;
        if (Value is { Length: > 0 }) count++;
        if (!string.IsNullOrWhiteSpace(Contains)) count++;
        if (!string.IsNullOrWhiteSpace(DoesNotContain)) count++;
        if (!string.IsNullOrWhiteSpace(StartsWith)) count++;
        if (!string.IsNullOrWhiteSpace(EndsWith)) count++;
        if (GreaterThanOrEqual.HasValue) count++;
        if (LessThanOrEqual.HasValue) count++;
        if (NotEqual.HasValue) count++;
        if (Between is { Length: > 0 }) count++;
        if (NotBetween is { Length: > 0 }) count++;

        if (count != 1)
        {
            throw new PSArgumentException("Specify exactly one AutoFilter condition.");
        }

        if (Between != null && Between.Length != 2)
        {
            throw new PSArgumentException("Between requires exactly two numeric values: minimum, maximum.");
        }

        if (NotBetween != null && NotBetween.Length != 2)
        {
            throw new PSArgumentException("NotBetween requires exactly two numeric values: minimum, maximum.");
        }
    }
}
