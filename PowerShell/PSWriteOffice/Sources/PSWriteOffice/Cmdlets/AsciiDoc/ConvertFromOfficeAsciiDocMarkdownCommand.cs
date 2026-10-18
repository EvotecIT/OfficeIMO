using System.IO;
using System.Management.Automation;
using OfficeIMO.AsciiDoc.Markdown;
using OfficeIMO.Markdown;

namespace PSWriteOffice.Cmdlets.AsciiDoc;

/// <summary>Converts Markdown to native AsciiDoc with fidelity diagnostics.</summary>
[Cmdlet(VerbsData.ConvertFrom, "OfficeAsciiDocMarkdown", DefaultParameterSetName = ParameterSetPath, SupportsShouldProcess = true)]
[OutputType(typeof(MarkdownToAsciiDocResult))]
public sealed class ConvertFromOfficeAsciiDocMarkdownCommand : PSCmdlet
{
    private const string ParameterSetPath = "Path";
    private const string ParameterSetDocument = "Document";

    /// <summary>Path to a Markdown file.</summary>
    [Parameter(Mandatory = true, Position = 0, ParameterSetName = ParameterSetPath)]
    public string Path { get; set; } = string.Empty;

    /// <summary>Markdown document to convert.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = ParameterSetDocument)]
    public MarkdownDoc Document { get; set; } = null!;

    /// <summary>Optional AsciiDoc destination path.</summary>
    [Parameter]
    public string? OutputPath { get; set; }

    /// <summary>Optional conversion settings.</summary>
    [Parameter]
    public MarkdownToAsciiDocOptions? Options { get; set; }

    /// <summary>Throw when a source feature cannot be mapped exactly.</summary>
    [Parameter]
    public SwitchParameter FailOnLoss { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        var document = ParameterSetName == ParameterSetPath
            ? MarkdownDoc.Load(SessionState.Path.GetUnresolvedProviderPathFromPSPath(Path))
            : Document;
        var result = document.ToAsciiDocDocumentResult(Options);
        if (FailOnLoss.IsPresent) result.RequireNoLoss();
        if (!string.IsNullOrWhiteSpace(OutputPath))
        {
            var output = SessionState.Path.GetUnresolvedProviderPathFromPSPath(OutputPath);
            if (!ShouldProcess(output, "Write converted AsciiDoc")) return;
            Directory.CreateDirectory(System.IO.Path.GetDirectoryName(output) ?? SessionState.Path.CurrentFileSystemLocation.Path);
            result.Value.Save(output);
        }
        WriteObject(result);
    }
}
