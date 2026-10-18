using System;
using System.Management.Automation;
using OfficeIMO.Excel;
using PSWriteOffice.Services.Excel;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Sets a built-in or application document property on an Excel workbook.</summary>
/// <example>
///   <summary>Stamp publishing metadata while creating a workbook.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>New-OfficeExcel -Path .\Report.xlsx {
///     Set-OfficeExcelDocumentProperty -Name Title -Value 'Operational dashboard'
///     Set-OfficeExcelDocumentProperty -Name Department -Value 'Operations' -Custom
/// }</code>
///   <para>Updates built-in or custom workbook properties through the current Excel DSL context.</para>
/// </example>
[Cmdlet(VerbsCommon.Set, "OfficeExcelDocumentProperty")]
[OutputType(typeof(ExcelDocument))]
public sealed class SetOfficeExcelDocumentPropertyCommand : PSCmdlet
{
    /// <summary>Workbook to update when provided explicitly.</summary>
    [Parameter(ValueFromPipeline = true)]
    public ExcelDocument? Document { get; set; }

    /// <summary>Property name to update.</summary>
    [Parameter(Mandatory = true, Position = 0)]
    public string Name { get; set; } = string.Empty;

    /// <summary>Property value.</summary>
    [Parameter(Position = 1)]
    public object? Value { get; set; }

    /// <summary>Emit the updated workbook.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <summary>Treat the property as a custom workbook property.</summary>
    [Parameter]
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
