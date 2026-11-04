using System;
using System.Collections;
using System.Collections.Generic;
using System.Management.Automation;
using OfficeIMO.Excel;
using PSWriteOffice.Foundation;
using PSWriteOffice.Services.Excel;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Sorts the used range on the current worksheet.</summary>
/// <example>
/// <summary>Sort by one column</summary>
/// <code>Open-OfficeExcel -Path .\Reports\Inventory.xlsx {
///     ExcelSheet 'Servers' {
///         ExcelSort Name
///     }
/// }</code>
/// <para>Sorts the used range of the Servers sheet by the Name header.</para>
/// </example>
/// <example>
/// <summary>Sort by several columns</summary>
/// <code>Open-OfficeExcel -Path .\Reports\Helpdesk.xlsx {
///     ExcelSheet 'Tickets' {
///         ExcelSort -Order ([ordered]@{ Priority = $true; AgeDays = $false })
///     }
/// }</code>
/// <para>Sorts tickets by priority ascending and then by age with the oldest first.</para>
/// </example>
[Cmdlet(VerbsLifecycle.Invoke, "OfficeExcelSort", DefaultParameterSetName = "Header")]
[Alias("ExcelSort")]
public sealed class InvokeOfficeExcelSortCommand : PSCmdlet
{

    /// <summary>Workbook to operate on outside the DSL context.</summary>
    [Parameter(ValueFromPipeline = true)]
    public ExcelDocument Document { get; set; } = null!;

    /// <summary>Worksheet name when using <see cref="Document"/>.</summary>
    [Parameter()]
    [OfficeSheetTransformation]
    public OfficeSheetSelector? Sheet { get; set; }

    /// <summary>Header to sort by (single-column sort).</summary>
    [Parameter(ParameterSetName = "Header", Mandatory = true, Position = 0)]
    public string Header { get; set; } = string.Empty;

    /// <summary>Sort descending (single-column sort).</summary>
    [Parameter(ParameterSetName = "Header")]
    public SwitchParameter Descending { get; set; }

    /// <summary>Ordered dictionary of header => ascending (true/false).</summary>
    [Parameter(ParameterSetName = "Order", Mandatory = true)]
    public System.Collections.Specialized.OrderedDictionary Order { get; set; } = new();

    /// <summary>Emit the worksheet after sorting.</summary>
    [Parameter()]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        var sheet = ResolveSheet();

        if (ParameterSetName == "Header")
        {
            if (string.IsNullOrWhiteSpace(Header))
            {
                throw new PSArgumentException("Header cannot be empty.");
            }

            sheet.SortUsedRangeByHeader(Header, ascending: !Descending.IsPresent);
        }
        else
        {
            var keys = ConvertOrder(Order);
            sheet.SortUsedRangeByHeaders(keys.ToArray());
        }

        if (PassThru.IsPresent)
        {
            WriteObject(sheet);
        }
    }

    private ExcelSheet ResolveSheet() => ExcelSheetResolver.ResolveTarget(this, Document, Sheet);

    private static List<(string Header, bool Ascending)> ConvertOrder(System.Collections.Specialized.OrderedDictionary order)
    {
        if (order == null || order.Count == 0)
        {
            throw new PSArgumentException("Order must contain at least one header.");
        }

        var list = new List<(string Header, bool Ascending)>();
        foreach (DictionaryEntry entry in order)
        {
            if (entry.Key == null) continue;
            var header = entry.Key.ToString();
            if (string.IsNullOrWhiteSpace(header)) continue;

            bool ascending = true;
            if (entry.Value != null)
            {
                switch (entry.Value)
                {
                    case bool b:
                        ascending = b;
                        break;
                    case string s:
                        ascending = !s.Equals("desc", StringComparison.OrdinalIgnoreCase)
                                    && !s.Equals("descending", StringComparison.OrdinalIgnoreCase)
                                    && !s.Equals("false", StringComparison.OrdinalIgnoreCase);
                        break;
                    default:
                        if (entry.Value is IConvertible convertible)
                        {
                            try
                            {
                                ascending = convertible.ToBoolean(System.Globalization.CultureInfo.InvariantCulture);
                            }
                            catch
                            {
                                ascending = true;
                            }
                        }
                        break;
                }
            }

            list.Add((header!, ascending));
        }

        if (list.Count == 0)
        {
            throw new PSArgumentException("Order must contain at least one header.");
        }

        return list;
    }
}
