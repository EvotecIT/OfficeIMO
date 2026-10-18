using System.IO;
using System.Management.Automation;
using OfficeIMO.Word;

namespace PSWriteOffice.Cmdlets.Word;

/// <summary>Compares two Word documents and optionally writes a tracked-change redline.</summary>
/// <example>
///   <summary>Create a structured comparison and redline.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$result = Compare-OfficeWordDocument -ReferencePath .\Before.docx -DifferencePath .\After.docx -RedlinePath .\Redline.docx</code>
///   <para>Returns deterministic findings and saves a Word document containing revision marks.</para>
/// </example>
/// <example>
///   <summary>Ignore whitespace and volatile metadata.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$options = New-OfficeWordComparisonOptions -IgnoreWhitespace -CompareVolatileMetadata:$false
/// Compare-OfficeWordDocument -ReferencePath .\Before.docx -DifferencePath .\After.docx -Options $options</code>
/// </example>
[Cmdlet(VerbsData.Compare, "OfficeWordDocument", SupportsShouldProcess = true)]
[OutputType(typeof(WordComparisonResult))]
public sealed class CompareOfficeWordDocumentCommand : PSCmdlet
{
    /// <summary>Path to the original Word document.</summary>
    [Parameter(Mandatory = true, Position = 0)]
    public string ReferencePath { get; set; } = string.Empty;

    /// <summary>Path to the modified Word document.</summary>
    [Parameter(Mandatory = true, Position = 1)]
    public string DifferencePath { get; set; } = string.Empty;

    /// <summary>Optional path for a tracked-change redline document.</summary>
    [Parameter]
    public string? RedlinePath { get; set; }

    /// <summary>Optional structural comparison switches.</summary>
    [Parameter]
    public WordComparisonOptions? Options { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        var reference = SessionState.Path.GetUnresolvedProviderPathFromPSPath(ReferencePath);
        var difference = SessionState.Path.GetUnresolvedProviderPathFromPSPath(DifferencePath);
        var result = WordDocumentComparer.CompareStructure(reference, difference, Options);
        if (!string.IsNullOrWhiteSpace(RedlinePath))
        {
            var output = SessionState.Path.GetUnresolvedProviderPathFromPSPath(RedlinePath!);
            if (ShouldProcess(output, "Write Word comparison redline"))
            {
                Directory.CreateDirectory(System.IO.Path.GetDirectoryName(output) ?? SessionState.Path.CurrentFileSystemLocation.Path);
                using var redline = WordDocumentComparer.Compare(reference, difference);
                redline.Save(output);
            }
        }
        WriteObject(result);
    }
}
