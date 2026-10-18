using System.IO;
using System.Management.Automation;
using OfficeIMO.AsciiDoc;

namespace PSWriteOffice.Cmdlets.AsciiDoc;

/// <summary>Saves an OfficeIMO AsciiDoc document.</summary>
/// <example>
///   <summary>Load, edit, and save an AsciiDoc document.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$document = Get-OfficeAsciiDoc -Path .\Guide.adoc
/// $document | Save-OfficeAsciiDoc -Path .\Guide-normalized.adoc -Mode Canonical</code>
/// </example>
[Cmdlet(VerbsData.Save, "OfficeAsciiDoc", SupportsShouldProcess = true)]
[OutputType(typeof(AsciiDocDocument))]
public sealed class SaveOfficeAsciiDocCommand : PSCmdlet
{
    /// <summary>AsciiDoc document to save.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, Position = 0)]
    public AsciiDocDocument Document { get; set; } = null!;

    /// <summary>Destination path.</summary>
    [Parameter(Mandatory = true, Position = 1)]
    public string Path { get; set; } = string.Empty;

    /// <summary>Optional writer settings.</summary>
    [Parameter]
    public AsciiDocWriterOptions? Options { get; set; }

    /// <summary>Writer mode. Preserve retains unchanged source; Canonical emits stable formatting.</summary>
    [Parameter]
    public AsciiDocWriterMode? Mode { get; set; }

    /// <summary>Canonical line ending: LF, CRLF, or CR. Omit it to retain the source preference.</summary>
    [Parameter]
    public OfficeLineEnding? LineEnding { get; set; }

    /// <summary>Return the saved document.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        var path = SessionState.Path.GetUnresolvedProviderPathFromPSPath(Path);
        if (!ShouldProcess(path, "Save AsciiDoc document")) return;
        Directory.CreateDirectory(System.IO.Path.GetDirectoryName(path) ?? SessionState.Path.CurrentFileSystemLocation.Path);
        Document.Save(path, BuildOptions());
        if (PassThru.IsPresent) WriteObject(Document);
    }

    private AsciiDocWriterOptions BuildOptions() {
        var options = new AsciiDocWriterOptions {
            Mode = Options?.Mode ?? AsciiDocWriterMode.Preserve,
            LineEnding = Options?.LineEnding
        };
        if (Mode.HasValue) options.Mode = Mode.Value;
        if (LineEnding.HasValue) options.LineEnding = OfficeLineEndingUtilities.ToText(LineEnding.Value);
        return options;
    }
}
