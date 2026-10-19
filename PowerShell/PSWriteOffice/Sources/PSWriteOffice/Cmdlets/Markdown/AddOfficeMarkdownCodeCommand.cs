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
///     Add-OfficeMarkdownCode -Language 'powershell' -Content $script
/// }</code>
///   <para>Appends a fenced code block with a language identifier.</para>
/// </example>
[Cmdlet(VerbsCommon.Add, "OfficeMarkdownCode", DefaultParameterSetName = ParameterSetContext)]
[Alias("MarkdownCode")]
[OutputType(typeof(MarkdownDoc))]
public sealed class AddOfficeMarkdownCodeCommand : PSCmdlet
{
    private const string ParameterSetContext = "Context";
    private const string ParameterSetDocument = "Document";

    /// <summary>Markdown document to update outside the DSL context.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = ParameterSetDocument)]
    public MarkdownDoc Document { get; set; } = null!;

    /// <summary>Code language identifier.</summary>
    [Parameter(Position = 0)]
    public string? Language { get; set; }

    /// <summary>Code content.</summary>
    [Parameter(Mandatory = true, Position = 1)]
    public string Content { get; set; } = string.Empty;

    /// <summary>Emit the Markdown document after appending the code block.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        var doc = ResolveDocument();
        doc.Code(Language ?? string.Empty, Content ?? string.Empty);

        if (PassThru.IsPresent)
        {
            WriteObject(doc);
        }
    }

    private MarkdownDoc ResolveDocument()
    {
        if (ParameterSetName == ParameterSetDocument)
        {
            return Document ?? throw new PSArgumentException("Provide a Markdown document.");
        }

        return MarkdownDslContext.Require(this).Document;
    }
}
