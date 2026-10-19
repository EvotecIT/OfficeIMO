using System;
using System.Management.Automation;
using OfficeIMO.Word;
using PSWriteOffice.Services.Word;

namespace PSWriteOffice.Cmdlets.Word;

/// <summary>Sets a built-in or custom document property on a Word document.</summary>
/// <example>
///   <summary>Set built-in and custom release metadata.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>New-OfficeWord -Path .\QuarterlyReport.docx {
///     Set-OfficeWordDocumentProperty -Name Title -Value 'Quarterly Report'
///     Set-OfficeWordDocumentProperty -Name ReleaseStatus -Value 'Approved' -Custom
///     Add-OfficeWordParagraph -Text 'Approved quarterly report'
/// }
/// Get-OfficeWordDocumentProperty -Path .\QuarterlyReport.docx -Name Title, ReleaseStatus</code>
///   <para>Writes document metadata during composition and reads it back for proof.</para>
/// </example>
[Cmdlet(VerbsCommon.Set, "OfficeWordDocumentProperty")]
[OutputType(typeof(WordDocument))]
public sealed class SetOfficeWordDocumentPropertyCommand : PSCmdlet
{
    /// <summary>Document to update when provided explicitly.</summary>
    [Parameter(ValueFromPipeline = true)]
    public WordDocument? Document { get; set; }

    /// <summary>Property name to update.</summary>
    [Parameter(Mandatory = true, Position = 0)]
    public string Name { get; set; } = string.Empty;

    /// <summary>Property value.</summary>
    [Parameter(Position = 1)]
    public object? Value { get; set; }

    /// <summary>Treat the property as a custom document property.</summary>
    [Parameter]
    public SwitchParameter Custom { get; set; }

    /// <summary>Emit the updated document.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        if (string.IsNullOrWhiteSpace(Name))
        {
            throw new PSArgumentException("Provide a document property name.", nameof(Name));
        }

        var document = Document ?? WordDslContext.Require(this).Document;
        if (document == null)
        {
            throw new InvalidOperationException("Word document was not provided.");
        }

        if (Custom.IsPresent)
        {
            WordDocumentPropertyService.SetCustomProperty(document, Name, Value);
        }
        else
        {
            WordDocumentPropertyService.SetBuiltInProperty(document, Name, Value);
        }

        if (PassThru.IsPresent)
        {
            WriteObject(document);
        }
    }
}
