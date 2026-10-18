using System;
using System.Collections;
using System.Collections.Generic;
using System.Management.Automation;
using OfficeIMO.Excel;
using PSWriteOffice.Services.Excel;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Sorts the used range on the current worksheet.</summary>
/// <example>
///   <summary>Sort by a single header.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>ExcelSheet 'Data' { Invoke-OfficeExcelSort -Header 'Name' }</code>
///   <para>Sorts by the Name column in ascending order.</para>
/// </example>
/// <example>
///   <summary>Sort by multiple headers in order.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$order = [ordered]@{ Status = $true; Total = $false }\nExcelSheet 'Data' { Invoke-OfficeExcelSort -Order $order }</code>
///   <para>Sorts by Status ascending, then Total descending.</para>
/// </example>
[Cmdlet(VerbsLifecycle.Invoke, "OfficeExcelSort", DefaultParameterSetName = ParameterSetContextSingle)]
[Alias("ExcelSort")]
public sealed class InvokeOfficeExcelSortCommand : PSCmdlet
{
    private const string ParameterSetContextSingle = "ContextSingle";
    private const string ParameterSetContextOrder = "ContextOrder";
    private const string ParameterSetDocumentSingle = "DocumentSingle";
    private const string ParameterSetDocumentOrder = "DocumentOrder";

    /// <summary>Workbook to operate on outside the DSL context.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = ParameterSetDocumentSingle)]
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = ParameterSetDocumentOrder)]
    public ExcelDocument Document { get; set; } = null!;

    /// <summary>Worksheet name when using <see cref="Document"/>.</summary>
    [Parameter(ParameterSetName = ParameterSetDocumentSingle)]
    [Parameter(ParameterSetName = ParameterSetDocumentOrder)]
    public string? Sheet { get; set; }

    /// <summary>Worksheet index (0-based) when using <see cref="Document"/>.</summary>
    [Parameter(ParameterSetName = ParameterSetDocumentSingle)]
    [Parameter(ParameterSetName = ParameterSetDocumentOrder)]
    public int? SheetIndex { get; set; }

    /// <summary>Header to sort by (single-column sort).</summary>
    [Parameter(Mandatory = true, Position = 0, ParameterSetName = ParameterSetContextSingle)]
    [Parameter(Mandatory = true, Position = 0, ParameterSetName = ParameterSetDocumentSingle)]
    public string Header { get; set; } = string.Empty;

    /// <summary>Sort descending (single-column sort).</summary>
    [Parameter(ParameterSetName = ParameterSetContextSingle)]
    [Parameter(ParameterSetName = ParameterSetDocumentSingle)]
    public SwitchParameter Descending { get; set; }

    /// <summary>Ordered dictionary of header => ascending (true/false).</summary>
    [Parameter(Mandatory = true, ParameterSetName = ParameterSetContextOrder)]
    [Parameter(Mandatory = true, ParameterSetName = ParameterSetDocumentOrder)]
    public Hashtable Order { get; set; } = new();

    /// <summary>Emit the worksheet after sorting.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        var sheet = ResolveSheet();

        if (ParameterSetName == ParameterSetContextSingle || ParameterSetName == ParameterSetDocumentSingle)
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

    private ExcelSheet ResolveSheet()
    {
        if (ParameterSetName == ParameterSetDocumentSingle || ParameterSetName == ParameterSetDocumentOrder)
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

    private static List<(string Header, bool Ascending)> ConvertOrder(Hashtable order)
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
