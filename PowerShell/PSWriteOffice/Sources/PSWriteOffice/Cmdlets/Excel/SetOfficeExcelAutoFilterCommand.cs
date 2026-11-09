using System.Collections;
using System.Globalization;
using System.Management.Automation;
using System.Security;
using OfficeIMO;
using OfficeIMO.Excel;
using PSWriteOffice.Foundation;
using PSWriteOffice.Services;
using PSWriteOffice.Services.Excel;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Adds an AutoFilter to a range and applies filter conditions.</summary>
/// <example>
/// <summary>Add filter buttons to a range</summary>
/// <code>Open-OfficeExcel -Path .\Reports\Inventory.xlsx {
///     ExcelSheet 'Servers' {
///         ExcelAutoFilter A1:F500
///     }
/// }</code>
/// <para>Adds AutoFilter drop-downs to the header row of A1:F500.</para>
/// </example>
/// <example>
/// <summary>Open a sheet filtered to the open work</summary>
/// <code>Open-OfficeExcel -Path .\Reports\Helpdesk.xlsx {
///     ExcelSheet 'Tickets' {
///         ExcelAutoFilter -Header Status -Value 'Open', 'On hold'
///         ExcelAutoFilter -Header AgeDays -GreaterThanOrEqual 3
///     }
/// }</code>
/// <para>Filters the ticket list to the Open and On hold statuses with at least three days of age, so the sheet opens on the work that needs attention.</para>
/// </example>
/// <example>
/// <summary>Filter a new range by header names</summary>
/// <code>Open-OfficeExcel -Path .\Reports\Helpdesk.xlsx {
///     ExcelSheet 'Tickets' {
///         ExcelAutoFilter A1:H900 -Criteria @{ Status = 'Open', 'On hold'; Region = 'EMEA' }
///     }
/// }</code>
/// <para>Adds filter buttons and immediately shows only open EMEA tickets.</para>
/// </example>
[Cmdlet(VerbsCommon.Set, "OfficeExcelAutoFilter", DefaultParameterSetName = "Range")]
[Alias("ExcelAutoFilter")]
public sealed class SetOfficeExcelAutoFilterCommand : PSCmdlet {
    /// <summary>Caller-owned workbook; omission uses the enclosing Excel content block.</summary>
    [Parameter(ValueFromPipeline = true)]
    public ExcelDocument? Document { get; set; }

    /// <summary>Worksheet name, one-based position or native worksheet object. Multi-sheet commands also accept wildcard names.</summary>
    [Parameter()]
    [OfficeSheetTransformation]
    public OfficeSheetSelector? Sheet { get; set; }

    /// <summary>Emits the native result after the operation.</summary>
    [Parameter()]
    public SwitchParameter PassThru { get; set; }

    /// <summary>A1 cell range on the selected worksheet.</summary>
    [Parameter(ParameterSetName = "Range", Mandatory = true, Position = 0)]
    [Parameter(ParameterSetName = "Condition")]
    public string Range { get; set; } = null!;

    /// <summary>Maps header names to filter values, value lists or comparison criteria.</summary>
    [Parameter(ParameterSetName = "Range")]
    public Hashtable? Criteria { get; set; }

    /// <summary>Header text identifying the target column.</summary>
    [Parameter(ParameterSetName = "Condition", Mandatory = true, Position = 0)]
    public string Header { get; set; } = null!;

    /// <summary>Value or values consumed by the selected operation.</summary>
    [Parameter(ParameterSetName = "Condition")]
    public string[]? Value { get; set; }

    /// <summary>Text that the column value must contain.</summary>
    [Parameter(ParameterSetName = "Condition")]
    public string? Contains { get; set; }

    /// <summary>Text that the column value must not contain.</summary>
    [Parameter(ParameterSetName = "Condition")]
    public string? DoesNotContain { get; set; }

    /// <summary>Text that the column value must start with.</summary>
    [Parameter(ParameterSetName = "Condition")]
    public string? StartsWith { get; set; }

    /// <summary>Text that the column value must end with.</summary>
    [Parameter(ParameterSetName = "Condition")]
    public string? EndsWith { get; set; }

    /// <summary>Numeric greater-than-or-equal condition.</summary>
    [Parameter(ParameterSetName = "Condition")]
    public double? GreaterThanOrEqual { get; set; }

    /// <summary>Numeric less-than-or-equal condition.</summary>
    [Parameter(ParameterSetName = "Condition")]
    public double? LessThanOrEqual { get; set; }

    /// <summary>Numeric not-equal condition.</summary>
    [Parameter(ParameterSetName = "Condition")]
    public double? NotEqual { get; set; }

    /// <summary>Inclusive numeric range condition. Provide exactly two values: minimum, maximum.</summary>
    [Parameter(ParameterSetName = "Condition")]
    public double[]? Between { get; set; }

    /// <summary>Outside numeric range condition. Provide exactly two values: minimum, maximum.</summary>
    [Parameter(ParameterSetName = "Condition")]
    public double[]? NotBetween { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        if (ParameterSetName == "Condition") ValidateSingleCondition();
        var sheet = ExcelSheetResolver.ResolveTarget(this, Document, Sheet);
        string target = Range ?? sheet.GetAutoFilters().FirstOrDefault(filter => !filter.IsTableFilter)?.Range ?? sheet.UsedRange.Address;
        var filter = sheet.Range(target).AutoFilter();
        if (ParameterSetName == "Range") {
            if (Criteria != null) foreach (DictionaryEntry item in Criteria) {
                object? value = OfficeArgumentValue.Unwrap(item.Value);
                var values = value is IEnumerable items && value is not string
                    ? items.Cast<object?>().Select(item => Convert.ToString(OfficeArgumentValue.Unwrap(item), CultureInfo.InvariantCulture) ?? string.Empty).ToArray()
                    : new[] { Convert.ToString(value, CultureInfo.InvariantCulture) ?? string.Empty };
                filter.Column(Convert.ToString(item.Key)!).Equals(values);
            }
        } else ApplyFilter(filter.Column(Header!));
        if (PassThru) WriteObject(sheet);
    }
    private void ApplyFilter(ExcelAutoFilterColumn column)
    {
        if (Value is { Length: > 0 })
        {
            column.Equals(Value);
        }
        else if (!string.IsNullOrWhiteSpace(Contains))
        {
            column.Contains(Contains!);
        }
        else if (!string.IsNullOrWhiteSpace(DoesNotContain))
        {
            column.DoesNotContain(DoesNotContain!);
        }
        else if (!string.IsNullOrWhiteSpace(StartsWith))
        {
            column.BeginsWith(StartsWith!);
        }
        else if (!string.IsNullOrWhiteSpace(EndsWith))
        {
            column.EndsWith(EndsWith!);
        }
        else if (GreaterThanOrEqual.HasValue)
        {
            column.GreaterThanOrEqual(GreaterThanOrEqual.Value);
        }
        else if (LessThanOrEqual.HasValue)
        {
            column.LessThanOrEqual(LessThanOrEqual.Value);
        }
        else if (NotEqual.HasValue)
        {
            column.NotEqual(NotEqual.Value);
        }
        else if (Between is { Length: 2 })
        {
            column.Between(Between[0], Between[1]);
        }
        else if (NotBetween is { Length: 2 })
        {
            column.NotBetween(NotBetween[0], NotBetween[1]);
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
