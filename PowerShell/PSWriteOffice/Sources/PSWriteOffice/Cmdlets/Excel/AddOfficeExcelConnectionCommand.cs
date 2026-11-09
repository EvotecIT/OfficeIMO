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

/// <summary>Adds Power Query or connection metadata that Excel can own and refresh.</summary>
/// <example>
/// <summary>Add a Power Query definition</summary>
/// <code>Open-OfficeExcel -Path .\Reports\Sales-2026.xlsx {
///     ExcelConnection -Name SalesQuery -CommandText 'let Source = Excel.CurrentWorkbook(){[Name="Orders"]}[Content] in Source'
/// }</code>
/// <para>Stores a Power Query query that Excel can refresh; OfficeIMO writes the definition and never runs it.</para>
/// </example>
/// <example>
/// <summary>Ship a workbook that refreshes from the warehouse</summary>
/// <code>$query = 'let Source = Sql.Database("dw.example.com", "Sales"), Orders = Source{[Schema="dbo", Item="Orders"]}[Data] in Orders'
/// Open-OfficeExcel -Path .\Reports\Live-Sales.xlsx {
///     ExcelConnection -Sheet Data -Name WarehouseOrders -CommandText $query -Description 'Orders from the sales warehouse' -QueryTable OrdersQuery -RefreshOnOpen
/// }</code>
/// <para>Adds a described query bound to a query table on the Data sheet and marks it to refresh when the workbook opens.</para>
/// </example>
[Cmdlet(VerbsCommon.Add, "OfficeExcelConnection", DefaultParameterSetName = "PowerQuery")]
[Alias("ExcelConnection")]
public sealed class AddOfficeExcelConnectionCommand : PSCmdlet {
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

    /// <summary>Connection name stored in workbook metadata.</summary>
    [Parameter(ParameterSetName = "PowerQuery", Mandatory = true)]
    public string Name { get; set; } = null!;

    /// <summary>Power Query M expression stored as metadata.</summary>
    [Parameter(ParameterSetName = "PowerQuery")]
    public string? CommandText { get; set; }

    /// <summary>Connection description.</summary>
    [Parameter(ParameterSetName = "PowerQuery")]
    public string? Description { get; set; }

    /// <summary>Native query-table metadata used by the connection.</summary>
    [Parameter(ParameterSetName = "PowerQuery")]
    public string? QueryTable { get; set; }

    /// <summary>Request refresh-on-open metadata.</summary>
    [Parameter(ParameterSetName = "PowerQuery")]
    public SwitchParameter RefreshOnOpen { get; set; }

    /// <summary>Metadata kind to add.</summary>
    [Parameter(ParameterSetName = "Xml", Mandatory = true)]
    public OfficeExcelPackageMetadataKind Kind { get; set; }

    /// <summary>XML payload to add as package metadata.</summary>
    [Parameter(ParameterSetName = "Xml", Mandatory = true)]
    public string Xml { get; set; } = null!;

    /// <inheritdoc />
    protected override void ProcessRecord() {
        var document = ExcelCommandTargets.Document(Document);
        ExcelSheet? sheet = Sheet == null ? (Document == null ? ExcelDslContext.Current?.CurrentSheet : null) : Sheet.Resolve(document.Sheets, s => s.Name);
        object result;
        if (ParameterSetName == "PowerQuery") result = document.AddPowerQueryMetadata(new ExcelPowerQueryMetadataOptions { Name = Name!, WorksheetName = sheet?.Name, QueryTableName = QueryTable, Description = Description, CommandText = CommandText, RefreshOnOpen = RefreshOnOpen });
        else result = Kind == OfficeExcelPackageMetadataKind.QueryTable ? document.AddWorksheetQueryTableMetadata((sheet ?? ExcelSheetResolver.Resolve(document)).Name, Xml!) : document.AddWorkbookConnectionMetadata(Xml!);
        if (PassThru) WriteObject(result);
    }

}
