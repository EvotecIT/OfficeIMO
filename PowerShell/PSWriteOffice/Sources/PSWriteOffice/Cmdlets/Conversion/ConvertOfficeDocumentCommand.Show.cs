using System.Management.Automation;
using System.Linq;
using OfficeIMO;
using PSWriteOffice.Services;

namespace PSWriteOffice.Cmdlets.Conversion;

public sealed partial class ConvertOfficeDocumentCommand {
    /// <summary>Open each newly published output in its default application. Requires file output; unavailable with AsString, Plan or SaveJob. Skipped and reused outputs are not opened.</summary>
    [Parameter(ParameterSetName = "Path")]
    [Parameter(ParameterSetName = "PathFile")]
    [Parameter(ParameterSetName = "LiteralPath")]
    [Parameter(ParameterSetName = "LiteralPathFile")]
    [Parameter(ParameterSetName = "Text")]
    [Parameter(ParameterSetName = "Document")]
    public SwitchParameter Show { get; set; }

    private void ValidateShow() {
        if (Show && (AsString || Plan || SaveJob != null ||
            (ParameterSetName is "Text" or "Document") && OutputPath == null))
            throw new PSArgumentException("Show requires file output and cannot be combined with AsString, Plan or SaveJob.");
    }

    private void OpenPublishedOutput(OfficeConversionReport report) {
        if (Show && !report.Diagnostics.Any(diagnostic => diagnostic.Severity == OfficeDiagnosticSeverity.Error))
            FileOpenService.Open(report.OutputPath!);
    }
}
