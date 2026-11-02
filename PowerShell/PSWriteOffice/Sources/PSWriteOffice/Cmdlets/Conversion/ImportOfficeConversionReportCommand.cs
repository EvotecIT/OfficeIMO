using OfficeIMO;
using System;
using System.Management.Automation;
using PSWriteOffice.Foundation;

namespace PSWriteOffice.Cmdlets.Conversion;

/// <summary>Reads saved conversion run reports without executing jobs or changing outputs.</summary>
/// <para>Reports describe an earlier attempt; checkpoint receipts and current source/output validation determine whether artifacts can be reused.</para>
/// <example><summary>Read the conversion runs in a directory.</summary>
/// <code>Import-OfficeConversionReport -Path ./runs/*.json | Select-Object State, Summary</code></example>
/// <example><summary>Read reports from FileInfo objects, preserving literal filenames.</summary>
/// <code>Get-ChildItem ./runs -Filter *.json -File | Import-OfficeConversionReport</code></example>
[Cmdlet(VerbsData.Import, "OfficeConversionReport", DefaultParameterSetName = "Path")]
[OutputType(typeof(OfficeBatchReport))]
public sealed class ImportOfficeConversionReportCommand : PSCmdlet {
    /// <summary>Saved run-report JSON paths, including provider wildcards.</summary>
    [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true, ParameterSetName = "Path")]
    [SupportsWildcards, ValidateNotNullOrEmpty]
    public string[] Path { get; set; } = Array.Empty<string>();

    /// <summary>Literal report paths, including FileInfo.FullName and provider PSPath pipeline properties.</summary>
    [Parameter(Mandatory = true, ValueFromPipelineByPropertyName = true, ParameterSetName = "LiteralPath")]
    [Alias("PSPath", "FullName"), ValidateNotNullOrEmpty]
    public string[] LiteralPath { get; set; } = Array.Empty<string>();

    /// <inheritdoc />
    protected override void ProcessRecord() {
        bool literal = ParameterSetName == "LiteralPath";
        foreach (string input in literal ? LiteralPath : Path) {
            if (Stopping) return;
            try {
                foreach (string path in OfficePath.ResolveInput(SessionState, input, literal, reportItemError:
                    (error, item) => WriteError(OfficeError.Create(error, "input-unreadable", item)))) {
                    if (Stopping) return;
                    try { WriteObject(OfficeBatchReport.Load(path)); }
                    catch (Exception error) when (error is not (ActionPreferenceStopException or PipelineStoppedException)) {
                        WriteError(OfficeError.Create(error, "input-unreadable", path));
                    }
                }
            } catch (Exception error) when (error is not (ActionPreferenceStopException or PipelineStoppedException)) {
                WriteError(OfficeError.Create(error, "input-unreadable", input));
            }
        }
    }
}
