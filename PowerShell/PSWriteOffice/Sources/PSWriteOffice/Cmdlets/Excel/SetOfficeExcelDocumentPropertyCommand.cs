using System.Management.Automation;
using OfficeIMO.Excel;
using PSWriteOffice.Services.Excel;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Sets a built-in or application document property on an Excel workbook.</summary>
/// <example>
/// <summary>Set the document title</summary>
/// <code>Open-OfficeExcel -Path .\Reports\Q3-Sales.xlsx {
///     Set-OfficeExcelDocumentProperty Title 'Q3 sales'
/// }</code>
/// <para>Sets the built-in Title property.</para>
/// </example>
/// <example>
/// <summary>Tag a workbook with custom metadata</summary>
/// <code>Open-OfficeExcel -Path .\Reports\Q3-Sales.xlsx {
///     Set-OfficeExcelDocumentProperty -Name Owner -Value 'Sales Operations' -Custom
///     Set-OfficeExcelDocumentProperty -Name Classification -Value 'Internal' -Custom
///     Set-OfficeExcelDocumentProperty -Name ReviewDate -Value (Get-Date).AddMonths(6) -Custom
/// }</code>
/// <para>Adds owner, classification and review-date custom properties that document management can search.</para>
/// </example>
[Cmdlet(VerbsCommon.Set, "OfficeExcelDocumentProperty", DefaultParameterSetName = "Default")]
[OutputType(typeof(ExcelDocument))]
public sealed class SetOfficeExcelDocumentPropertyCommand : PSCmdlet
{
    /// <summary>Workbook to update when provided explicitly.</summary>
    [Parameter(ParameterSetName = "Default", ValueFromPipeline = true)]
    public ExcelDocument? Document { get; set; }

    /// <summary>Property name to update.</summary>
    [Parameter(ParameterSetName = "Default", Mandatory = true, Position = 0)]
    public string Name { get; set; } = string.Empty;

    /// <summary>Property value.</summary>
    [Parameter(ParameterSetName = "Default", Position = 1)]
    public object? Value { get; set; }

    /// <summary>Emit the updated workbook.</summary>
    [Parameter(ParameterSetName = "Default")]
    public SwitchParameter PassThru { get; set; }

    /// <summary>Treat the property as a custom workbook property.</summary>
    [Parameter(ParameterSetName = "Default")]
    public SwitchParameter Custom { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        if (string.IsNullOrWhiteSpace(Name))
        {
            throw new PSArgumentException("Provide a document property name.", nameof(Name));
        }

        var document = Document ?? ExcelDslContext.Require(this).Document;
        if (document == null)
        {
            throw new InvalidOperationException("Excel workbook was not provided.");
        }

        ExcelDocumentPropertyService.SetProperty(document, Name, Value, Custom.IsPresent);

        if (PassThru.IsPresent)
        {
            WriteObject(document);
        }
    }
}
