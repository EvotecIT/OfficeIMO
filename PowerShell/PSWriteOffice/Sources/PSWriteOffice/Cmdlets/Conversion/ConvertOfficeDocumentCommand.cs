using System;
using System.Collections.Generic;
using System.IO;
using System.Management.Automation;
using System.Security;
using System.Threading.Tasks;
using OfficeIMO;
using PSWriteOffice.Foundation;
using PSWriteOffice.Services;

namespace PSWriteOffice.Cmdlets.Conversion;

/// <summary>Converts files through OfficeIMO's native routes and portable batch engine.</summary>
/// <para>Writes each completed file's conversion evidence as it finishes. Paths resolve through the FileSystem provider before background execution; LiteralPath preserves wildcard characters. Saved jobs use the same contract as Studio and the CLI.</para>
/// <example>
///   <summary>Convert one document to PDF.</summary>
///   <code>Convert-OfficeDocument -Path ./report.docx -OutputPath ./report.pdf</code>
/// </example>
/// <example>
///   <summary>Convert a file pipeline and retain complete batch counts.</summary>
///   <code>Get-ChildItem ./reports -Filter *.docx | Convert-OfficeDocument -To Pdf -OutputDirectory ./pdf -SummaryVariable summary
/// $summary</code>
/// </example>
/// <example>
///   <summary>Save and preview a portable conversion job.</summary>
///   <code>Convert-OfficeDocument -Path ./reports -Recurse -To Pdf -OutputDirectory ./pdf -SaveJob ./reports.officejob.json
/// Convert-OfficeDocument -JobPath ./reports.officejob.json -Plan</code>
/// </example>
[Cmdlet("Convert", "OfficeDocument", DefaultParameterSetName = "Path", SupportsShouldProcess = true)]
[OutputType(typeof(OfficeConversionReport), typeof(OfficeBatchItemReport), typeof(OfficeBatchPlan), typeof(FileInfo), typeof(string))]
public sealed partial class ConvertOfficeDocumentCommand : AsyncPSCmdlet {
    private readonly List<string> _inputs = new();

    /// <summary>Input files or folders, with FileSystem provider wildcard expansion.</summary>
    [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true, ParameterSetName = "Path")]
    [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true, ParameterSetName = "PathFile")]
    [SupportsWildcards]
    [ValidateNotNullOrEmpty]
    public string[] Path { get; set; } = Array.Empty<string>();

    /// <summary>Input files or folders without wildcard expansion; accepts FileInfo.FullName and provider PSPath.</summary>
    [Parameter(Mandatory = true, ValueFromPipelineByPropertyName = true, ParameterSetName = "LiteralPath")]
    [Parameter(Mandatory = true, ValueFromPipelineByPropertyName = true, ParameterSetName = "LiteralPathFile")]
    [Alias("FullName", "PSPath")]
    [ValidateNotNullOrEmpty]
    public string[] LiteralPath { get; set; } = Array.Empty<string>();

    /// <summary>Single-file destination. Its extension selects the target when To is omitted.</summary>
    [Parameter(Mandatory = true, Position = 1, ParameterSetName = "PathFile")]
    [Parameter(Mandatory = true, Position = 1, ParameterSetName = "LiteralPathFile")]
    [Parameter(ParameterSetName = "Text")]
    [Parameter(ParameterSetName = "Document")]
    public string? OutputPath { get; set; }

    /// <summary>Destination folder for a batch.</summary>
    [Parameter(ParameterSetName = "Path")]
    [Parameter(ParameterSetName = "LiteralPath")]
    public string? OutputDirectory { get; set; }

    /// <summary>Target format family, such as Pdf or Word, or an exact physical extension, such as .doc or .docx.</summary>
    [Parameter(Mandatory = true, ParameterSetName = "Path")]
    [Parameter(Mandatory = true, ParameterSetName = "LiteralPath")]
    [Parameter(ParameterSetName = "PathFile")]
    [Parameter(ParameterSetName = "LiteralPathFile")]
    [Parameter(Mandatory = true, ParameterSetName = "Text")]
    [Parameter(Mandatory = true, ParameterSetName = "Document")]
    [ArgumentCompleter(typeof(OfficeFormatCompleter))]
    public string? To { get; set; }

    /// <summary>Format-owned settings, typed conversion options, or property hashtables.</summary>
    [Parameter(ParameterSetName = "Path")]
    [Parameter(ParameterSetName = "LiteralPath")]
    [Parameter(ParameterSetName = "PathFile")]
    [Parameter(ParameterSetName = "LiteralPathFile")]
    [Parameter(ParameterSetName = "Text")]
    [Parameter(ParameterSetName = "Document")]
    public object[]? Options { get; set; }

    /// <summary>Runtime source password; saved definitions contain references rather than password values.</summary>
    [Parameter(ParameterSetName = "Path")]
    [Parameter(ParameterSetName = "LiteralPath")]
    [Parameter(ParameterSetName = "PathFile")]
    [Parameter(ParameterSetName = "LiteralPathFile")]
    [Parameter(ParameterSetName = "Text")]
    [Parameter(ParameterSetName = "Document")]
    [OfficePasswordTransformation]
    public SecureString? Password { get; set; }

    /// <summary>Exact executable route identifier from Get-OfficeDocumentCapability -Conversion.</summary>
    [Parameter(ParameterSetName = "Path")]
    [Parameter(ParameterSetName = "LiteralPath")]
    [Parameter(ParameterSetName = "PathFile")]
    [Parameter(ParameterSetName = "LiteralPathFile")]
    [Parameter(ParameterSetName = "Text")]
    [Parameter(ParameterSetName = "Document")]
    public string? Route { get; set; }

    /// <summary>Output intent, subject to the selected route's supported profiles.</summary>
    [Parameter(ParameterSetName = "Path")]
    [Parameter(ParameterSetName = "LiteralPath")]
    [Parameter(ParameterSetName = "PathFile")]
    [Parameter(ParameterSetName = "LiteralPathFile")]
    [Parameter(ParameterSetName = "Text")]
    [Parameter(ParameterSetName = "Document")]
    public OfficeConversionProfile OutputProfile { get; set; } = OfficeConversionProfile.Faithful;

    /// <summary>Reject a conversion that reports fidelity loss before publishing its output.</summary>
    [Parameter(ParameterSetName = "Path")]
    [Parameter(ParameterSetName = "LiteralPath")]
    [Parameter(ParameterSetName = "PathFile")]
    [Parameter(ParameterSetName = "LiteralPathFile")]
    [Parameter(ParameterSetName = "Text")]
    [Parameter(ParameterSetName = "Document")]
    public SwitchParameter FailOnLoss { get; set; }

    /// <summary>Explicitly accept the native legacy Word importer's reported reductions.</summary>
    [Parameter(ParameterSetName = "Path")]
    [Parameter(ParameterSetName = "LiteralPath")]
    [Parameter(ParameterSetName = "PathFile")]
    [Parameter(ParameterSetName = "LiteralPathFile")]
    [Parameter(ParameterSetName = "Text")]
    [Parameter(ParameterSetName = "Document")]
    public SwitchParameter AllowLegacyImportLoss { get; set; }

    /// <summary>Apply verified lossless compression to PDF output.</summary>
    [Parameter(ParameterSetName = "Path")]
    [Parameter(ParameterSetName = "LiteralPath")]
    [Parameter(ParameterSetName = "PathFile")]
    [Parameter(ParameterSetName = "LiteralPathFile")]
    [Parameter(ParameterSetName = "Text")]
    [Parameter(ParameterSetName = "Document")]
    public SwitchParameter CompressPdf { get; set; }

    /// <summary>Leave an existing destination unchanged and report its input as skipped.</summary>
    [Parameter(ParameterSetName = "Path")]
    [Parameter(ParameterSetName = "LiteralPath")]
    [Parameter(ParameterSetName = "PathFile")]
    [Parameter(ParameterSetName = "LiteralPathFile")]
    [Parameter(ParameterSetName = "Text")]
    [Parameter(ParameterSetName = "Document")]
    public SwitchParameter NoClobber { get; set; }

    /// <summary>Skip a file whose existing output has a strictly newer UTC modification time.</summary>
    [Parameter(ParameterSetName = "Path")]
    [Parameter(ParameterSetName = "LiteralPath")]
    [Parameter(ParameterSetName = "PathFile")]
    [Parameter(ParameterSetName = "LiteralPathFile")]
    public SwitchParameter SkipUpToDate { get; set; }

    /// <summary>Choose a unique destination name when the requested output exists.</summary>
    [Parameter(ParameterSetName = "Path")]
    [Parameter(ParameterSetName = "LiteralPath")]
    [Parameter(ParameterSetName = "PathFile")]
    [Parameter(ParameterSetName = "LiteralPathFile")]
    [Parameter(ParameterSetName = "Text")]
    [Parameter(ParameterSetName = "Document")]
    public SwitchParameter AutoRename { get; set; }

    /// <summary>Maximum source bytes per file.</summary>
    [Parameter(ParameterSetName = "Path")]
    [Parameter(ParameterSetName = "LiteralPath")]
    [Parameter(ParameterSetName = "PathFile")]
    [Parameter(ParameterSetName = "LiteralPathFile")]
    [Parameter(ParameterSetName = "Text")]
    [Parameter(ParameterSetName = "Document")]
    [ValidateRange(1L, long.MaxValue)]
    public long MaximumInputBytes { get; set; } = 64L * 1024 * 1024;

    /// <summary>Maximum serialized output bytes per file.</summary>
    [Parameter(ParameterSetName = "Path")]
    [Parameter(ParameterSetName = "LiteralPath")]
    [Parameter(ParameterSetName = "PathFile")]
    [Parameter(ParameterSetName = "LiteralPathFile")]
    [Parameter(ParameterSetName = "Text")]
    [Parameter(ParameterSetName = "Document")]
    [ValidateRange(1L, long.MaxValue)]
    public long MaximumOutputBytes { get; set; } = 256L * 1024 * 1024;

    /// <summary>Maximum XML characters in one package part.</summary>
    [Parameter(ParameterSetName = "Path")]
    [Parameter(ParameterSetName = "LiteralPath")]
    [Parameter(ParameterSetName = "PathFile")]
    [Parameter(ParameterSetName = "LiteralPathFile")]
    [Parameter(ParameterSetName = "Text")]
    [Parameter(ParameterSetName = "Document")]
    [ValidateRange(1L, long.MaxValue)]
    public long MaximumXmlCharactersInPart { get; set; } = 10L * 1024 * 1024;

    /// <summary>Final batch counts or single-file evidence assigned before any terminal batch error.</summary>
    [Parameter]
    public string? SummaryVariable { get; set; }

    /// <inheritdoc />
    protected override Task BeginProcessingAsync() {
        OfficeWorkflowCommandUtilities.RegisterConversions();
        return Task.CompletedTask;
    }

    /// <inheritdoc />
    protected override async Task ProcessRecordAsync() {
        if (ParameterSetName == "Job") return;
        if (ParameterSetName == "Text") { await ConvertTextAsync(); return; }
        if (ParameterSetName == "Document") {
            if (OfficePath.GetPipelineLiteralPaths(Document) is { } paths)
                foreach (string path in paths) AddInput(path, literal: true);
            else await ConvertLiveDocumentAsync();
            return;
        }
        bool literal = ParameterSetName.StartsWith("LiteralPath", StringComparison.Ordinal);
        foreach (string path in literal ? LiteralPath : Path) AddInput(path, literal);
    }

    private void AddInput(string path, bool literal) {
        foreach (string input in OfficePath.ResolveInput(SessionState, path, literal, allowDirectory: !IsSingleFile)) {
            int limit = IsSingleFile ? 1 : MaximumFiles == 0 ? int.MaxValue : MaximumFiles;
            if (_inputs.Count >= limit) throw new PSArgumentException("The input selection exceeds MaximumFiles or the single-file destination limit.");
            _inputs.Add(input);
        }
    }

    private bool IsSingleFile => ParameterSetName.EndsWith("File", StringComparison.Ordinal) ||
        ParameterSetName == "Document" && OutputPath != null;

    /// <inheritdoc />
    protected override async Task EndProcessingAsync() {
        if (ParameterSetName == "Text" || ParameterSetName == "Document" && _inputs.Count == 0) return;
        if (NoClobber && AutoRename) throw new PSArgumentException("NoClobber and AutoRename cannot be combined.");
        if (IsSingleFile) await ConvertFileAsync();
        else if (AsString) await ConvertFilesToTextAsync();
        else await ConvertBatchAsync();
    }

    private OfficePublicationPolicy ConflictPolicy => NoClobber ? OfficePublicationPolicy.Fail :
        AutoRename ? OfficePublicationPolicy.Rename : OfficePublicationPolicy.Replace;

    private void SetSummary(object value) {
        if (!string.IsNullOrWhiteSpace(SummaryVariable)) SessionState.PSVariable.Set(SummaryVariable, value);
    }

    private async Task ConvertFileAsync() {
        if (_inputs.Count != 1) throw new PSArgumentException("Select exactly one file with OutputPath.");
        string output = OfficePath.ResolveOutput(SessionState, OutputPath!, directory: false);
        if (OfficeFormatSelection.Parse(To) is { } target && !target.Matches(System.IO.Path.GetExtension(output)))
            throw new PSArgumentException("To and the OutputPath extension must name the same target format.");
        var options = CaptureOptions(_inputs, System.IO.Path.GetExtension(output));
        if (!ShouldProcess(output, "Convert document")) return;
        if (OfficeBatchOutputSelection.ObserveSkip(_inputs[0], output, NoClobber.IsPresent, SkipUpToDate.IsPresent) is { } skipped) {
            SetSummary(skipped);
            WriteObject(skipped);
            return;
        }
        var result = await OfficeConverter.ConvertAsync(_inputs[0], output, options, CancelToken);
        SetSummary(result);
        WriteObject(result);
    }
}
