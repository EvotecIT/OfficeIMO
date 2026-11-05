using System.Management.Automation;
using OfficeIMO.Markdown;
using PSWriteOffice.Services.Markdown;

namespace PSWriteOffice.Cmdlets.Markdown;

/// <summary>Adds a Markdown code block.</summary>
/// <example>
///   <summary>Add a PowerShell command block to documentation.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$script = @(
///     'Import-Module PSWriteOffice'
///     'Get-OfficeExcelSummary -Path .\report.xlsx'
/// ) -join [Environment]::NewLine
/// New-OfficeMarkdown -Path .\Cookbook.md {
///     Add-OfficeMarkdownHeading -Level 2 -Text 'Workbook summary'
///     Add-OfficeMarkdownCode -Language 'powershell' -Text $script
/// }</code>
///   <para>Appends a fenced code block with a language identifier.</para>
/// </example>
[Cmdlet(VerbsCommon.Add, "OfficeMarkdownCode", DefaultParameterSetName = ParameterSetContext)]
[OutputType(typeof(MarkdownDocument))]
[Alias("MarkdownCode") ]
public sealed class AddOfficeMarkdownCodeCommand : PSCmdlet
{
    private const string ParameterSetContext = "Context";
    private const string ParameterSetDocument = "Document";

    /// <summary>Markdown document to update outside the DSL context.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = ParameterSetDocument)]
    public MarkdownDocument Document { get; set; } = null!;

    /// <summary>Code language identifier.</summary>
    [Parameter(Position = 0)]
    public string? Language { get; set; }

    /// <summary>Code content.</summary>
    [Parameter(Mandatory = true, Position = 1)]
    public string Text { get; set; } = string.Empty;

    /// <summary>Emit the Markdown document after appending the code block.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        var doc = ResolveDocument();
        doc.Code(Language ?? string.Empty, Text ?? string.Empty);

        if (PassThru.IsPresent)
        {
            WriteObject(doc);
        }
    }

    private MarkdownDocument ResolveDocument()
    {
        if (ParameterSetName == ParameterSetDocument)
        {
            return Document ?? throw new PSArgumentException("Provide a Markdown document.");
        }

        return MarkdownDslContext.Require(this).Document;
    }
}
