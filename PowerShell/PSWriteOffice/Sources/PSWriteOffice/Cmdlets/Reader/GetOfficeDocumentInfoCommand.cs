using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Management.Automation;
using System.Net;
using System.Security;
using System.Threading.Tasks;
using OfficeIMO;
using OfficeIMO.Excel;
using OfficeIMO.PowerPoint;
using OfficeIMO.Reader;
using PSWriteOffice.Foundation;
using PSWriteOffice.Services.Reader;

namespace PSWriteOffice.Cmdlets.Reader;

/// <summary>Inspects format, native counts and encryption without changing a source or reopening a live document.</summary>
/// <example><summary>Inventory workbook counts and protection.</summary>
/// <code>Get-ChildItem ./Reports/*.xlsx | Get-OfficeDocumentInfo | Select-Object Path, Format, SheetCount, IsEncrypted</code></example>
/// <example><summary>Inspect unsaved presentation edits.</summary>
/// <code>$presentation | Get-OfficeDocumentInfo -Detailed</code></example>
[Cmdlet(VerbsCommon.Get, "OfficeDocumentInfo", DefaultParameterSetName = "Path")]
[OutputType(typeof(OfficeDocumentInfo))]
public sealed class GetOfficeDocumentInfoCommand : OfficeInputCmdlet {
    private OfficeInspectOptions _options = null!;

    /// <summary>Live native model or normalized Reader result. Its source path is never reopened.</summary>
    [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true, ParameterSetName = "Document")]
    public object Document { get; set; } = null!;
    /// <summary>Include the native format's detached inspection snapshot.</summary>
    [Parameter] public SwitchParameter Detailed { get; set; }
    /// <summary>Inspection, native format or Reader policy objects; hashtables are accepted.</summary>
    [Parameter] [OfficeInspectionOptionsTransformation] public object[]? Options { get; set; }
    /// <summary>Password used locally to open an encrypted source.</summary>
    [Parameter(ParameterSetName = "Path")] [Parameter(ParameterSetName = "LiteralPath")]
    [OfficePasswordTransformation] public SecureString? Password { get; set; }
    /// <summary>Content and filename selection policy.</summary>
    [Parameter(ParameterSetName = "Path")] [Parameter(ParameterSetName = "LiteralPath")]
    public ReaderDetectionMode? DetectionMode { get; set; }
    /// <summary>Maximum prefix bytes examined for content signatures.</summary>
    [Parameter(ParameterSetName = "Path")] [Parameter(ParameterSetName = "LiteralPath")]
    [ValidateRange(1, int.MaxValue)] public int? MaximumProbeBytes { get; set; }
    /// <summary>Maximum container entries examined during format detection.</summary>
    [Parameter(ParameterSetName = "Path")] [Parameter(ParameterSetName = "LiteralPath")]
    [ValidateRange(1, int.MaxValue)] public int? MaximumContainerEntries { get; set; }
    /// <summary>Skip structural container inspection when identifying the format.</summary>
    [Parameter(ParameterSetName = "Path")] [Parameter(ParameterSetName = "LiteralPath")]
    public SwitchParameter NoContainerInspection { get; set; }
    /// <summary>Immutable caller-configured format registry.</summary>
    [Parameter(ParameterSetName = "Path")] [Parameter(ParameterSetName = "LiteralPath")]
    public OfficeDocumentReader? Reader { get; set; }

    /// <inheritdoc />
    protected override Task BeginProcessingAsync() {
        _options = OfficeInspectionConfiguration.Bind(Options, Reader);
        _options.IncludeSummary = true;
        _options.IncludeContent = false;
        _options.Detailed = MyInvocation.BoundParameters.ContainsKey(nameof(Detailed)) ? Detailed.IsPresent : _options.Detailed;
        _options.Detection = _options.Detection == null ? new ReaderDetectionOptions {
            Mode = _options.Reader.DetectionMode, MaxProbeBytes = _options.Reader.DetectionMaxProbeBytes,
            MaxContainerEntries = _options.Reader.DetectionMaxContainerEntries,
            InspectContainers = _options.Reader.DetectionInspectContainers
        } : OfficeReadConfiguration.Copy(_options.Detection);
        if (DetectionMode.HasValue) _options.Detection.Mode = DetectionMode.Value;
        if (MaximumProbeBytes.HasValue) _options.Detection.MaxProbeBytes = MaximumProbeBytes.Value;
        if (MaximumContainerEntries.HasValue) _options.Detection.MaxContainerEntries = MaximumContainerEntries.Value;
        if (NoContainerInspection) _options.Detection.InspectContainers = false;
        if (MyInvocation.BoundParameters.ContainsKey(nameof(Password)))
            _options.Reader.OpenPassword = Password == null ? null : new NetworkCredential(string.Empty, Password).Password;
        return Task.CompletedTask;
    }

    /// <inheritdoc />
    protected override async Task ProcessRecordAsync() {
        if (ParameterSetName == "Document") {
            object value = OfficeArgumentValue.Unwrap(Document)!;
            // An object parameter can bind a FileInfo before property-name binding selects LiteralPath.
            // Resolve that filesystem input on the pipeline thread, just like the inherited path sets.
            if (value is FileInfo file) {
                IReadOnlyList<string> paths = OfficePath.ResolveInput(SessionState, file.FullName, true, false,
                    (error, input) => WriteError(OfficeError.Create(error, "input-unreadable", input)));
                await RunBatchAsync(paths, (path, token) => Task.FromResult<object?>(OfficeDocument.Inspect(path, _options, token)), "Inspecting documents");
            } else {
                try { WriteObject(OfficeDocument.Inspect(value, _options, CancelToken)); }
                catch (Exception error) when (error is not (OperationCanceledException or PipelineStoppedException or ActionPreferenceStopException)) {
                    WriteError(OfficeError.Create(error, "inspection-failed", value));
                }
            }
            return;
        }
        IReadOnlyList<string> inputs = ResolveInputs();
        await RunBatchAsync(inputs, (path, token) => Task.FromResult<object?>(OfficeDocument.Inspect(path, _options, token)), "Inspecting documents");
    }
}
