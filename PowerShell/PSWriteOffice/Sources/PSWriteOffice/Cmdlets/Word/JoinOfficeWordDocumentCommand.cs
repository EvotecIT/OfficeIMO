using System;
using System.IO;
using System.Management.Automation;
using OfficeIMO.Word;
using PSWriteOffice.Services;
using PSWriteOffice.Services.Word;

namespace PSWriteOffice.Cmdlets.Word;

/// <summary>Appends one or more Word documents into a base Word document.</summary>
/// <para>Uses OfficeIMO.Word document append support and preserves the wrapper as an operator-friendly merge command.</para>
/// <example>
///   <summary>Merge a cover, body, and appendix into a release packet.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$proof = @(
///     Join-OfficeWordDocument -Path .\Cover.docx -AppendPath .\Body.docx, .\Appendix.docx -OutputPath .\ReleasePacket.docx
///     Get-OfficeWordStatistics -Path .\ReleasePacket.docx |
///         Select-Object -Property Paragraphs, Tables, Images
/// )
/// $proof</code>
///   <para>Appends the source documents with OfficeIMO.Word and then reads back basic structure from the merged output.</para>
/// </example>
/// <example>
///   <summary>Append into an already-open document object.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$doc = Get-OfficeWord -Path .\Base.docx
/// $doc | Join-OfficeWordDocument -AppendPath .\Section1.docx, .\Section2.docx -PassThru |
///     Save-OfficeWord -Path .\Combined.docx</code>
///   <para>Keeps the wrapper thin by piping the OfficeIMO document object through append and save commands.</para>
/// </example>
[Cmdlet(VerbsCommon.Join, "OfficeWordDocument", DefaultParameterSetName = ParameterSetPath, SupportsShouldProcess = true)]
[Alias("Merge-OfficeWordDocument", "WordDocumentJoin")]
[OutputType(typeof(WordDocument))]
public sealed class JoinOfficeWordDocumentCommand : PSCmdlet {
    private const string ParameterSetPath = "Path";
    private const string ParameterSetDocument = "Document";

    /// <summary>Base document path.</summary>
    [Parameter(Mandatory = true, Position = 0, ParameterSetName = ParameterSetPath)]
    [Alias("InputPath", "BasePath")]
    public string Path { get; set; } = string.Empty;

    /// <summary>Base document object.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = ParameterSetDocument)]
    public WordDocument Document { get; set; } = null!;

    /// <summary>Documents to append to the base document.</summary>
    [Parameter(Mandatory = true, Position = 1)]
    [Alias("SourcePath")]
    public string[] AppendPath { get; set; } = Array.Empty<string>();

    /// <summary>Optional output path. When omitted for path input, the base document is updated in place.</summary>
    [Parameter]
    public string? OutputPath { get; set; }

    /// <summary>Open the saved output with the shell.</summary>
    [Parameter]
    [Alias("Show")]
    public SwitchParameter Open { get; set; }

    /// <summary>Emit the merged Word document instead of disposing it.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        WordDocument? document = null;
        var dispose = false;

        try {
            string? saveTarget = null;
            if (ParameterSetName == ParameterSetPath) {
                var resolvedPath = ResolveExistingPath(Path);
                document = WordDocumentService.LoadDocument(resolvedPath, readOnly: false, autoSave: false);
                dispose = true;
                saveTarget = string.IsNullOrWhiteSpace(OutputPath) ? resolvedPath : ResolveOutputPath(OutputPath!);
            } else {
                document = Document;
                if (!string.IsNullOrWhiteSpace(OutputPath)) {
                    saveTarget = ResolveOutputPath(OutputPath!);
                } else if (Open.IsPresent) {
                    saveTarget = document.FilePath ?? throw new InvalidOperationException("No saved file path was available.");
                }
            }

            var processTarget = !string.IsNullOrWhiteSpace(saveTarget)
                ? saveTarget!
                : document.FilePath ?? "Word document";
            var processAction = !string.IsNullOrWhiteSpace(saveTarget)
                ? "Write joined Word document"
                : "Join Word documents";
            if (!ShouldProcess(processTarget, processAction)) {
                return;
            }

            foreach (var sourcePath in AppendPath) {
                using var source = WordDocumentService.LoadDocument(ResolveExistingPath(sourcePath), readOnly: true, autoSave: false);
                document.AppendDocument(source);
            }

            string? savedPath = null;

            if (!string.IsNullOrWhiteSpace(OutputPath)) {
                savedPath = saveTarget;
                document.Save(savedPath!);
            } else if (ParameterSetName == ParameterSetPath) {
                document.Save();
                savedPath = saveTarget;
            } else if (Open.IsPresent) {
                savedPath = saveTarget;
                document.Save();
            }

            if (Open.IsPresent) {
                FileOpenService.Open(savedPath ?? throw new InvalidOperationException("No saved file path was available."));
            }

            if (PassThru.IsPresent) {
                dispose = false;
                WriteObject(document);
            }
        } finally {
            if (dispose) {
                document?.Dispose();
            }
        }
    }

    private string ResolveExistingPath(string path) {
        var resolvedPath = SessionState.Path.GetUnresolvedProviderPathFromPSPath(path);
        if (!File.Exists(resolvedPath)) {
            throw new FileNotFoundException($"File {resolvedPath} doesn't exist.", resolvedPath);
        }
        return resolvedPath;
    }

    private string ResolveOutputPath(string path) {
        var providerPath = SessionState.Path.GetUnresolvedProviderPathFromPSPath(path);
        return System.IO.Path.IsPathRooted(providerPath)
            ? providerPath
            : System.IO.Path.Combine(SessionState.Path.CurrentFileSystemLocation.Path, providerPath);
    }
}