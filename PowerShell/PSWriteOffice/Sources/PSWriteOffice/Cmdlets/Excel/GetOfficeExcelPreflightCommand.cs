using System;
using System.Linq;
using System.Management.Automation;
using OfficeIMO.Excel;
using PSWriteOffice.Services.Excel;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Runs OfficeIMO Excel feature and workflow preflight checks.</summary>
/// <example>
///   <summary>Check workbook edit and export readiness.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$preflight = Get-OfficeExcelPreflight -Path .\Report.xlsx -Capability EditCellValues,ExportPdfReport -IncludeFeatures -IncludeRepairHints
/// $preflight.Capabilities |
///     Where-Object Passed -eq $false</code>
///   <para>Returns reusable OfficeIMO capability diagnostics and discovered workbook features.</para>
/// </example>
[Cmdlet(VerbsCommon.Get, "OfficeExcelPreflight", DefaultParameterSetName = ParameterSetPath)]
[Alias("ExcelPreflight")]
[OutputType(typeof(PSObject), typeof(string))]
public sealed class GetOfficeExcelPreflightCommand : PSCmdlet {
    private const string ParameterSetPath = "Path";
    private const string ParameterSetDocument = "Document";

    /// <summary>Path to the workbook.</summary>
    [Parameter(Mandatory = true, Position = 0, ParameterSetName = ParameterSetPath)]
    [Alias("InputPath", "FilePath")]
    public string Path { get; set; } = string.Empty;

    /// <summary>Workbook to inspect.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = ParameterSetDocument)]
    public ExcelDocument Document { get; set; } = null!;

    /// <summary>Capabilities to evaluate or enforce. Defaults to all capabilities.</summary>
    [Parameter]
    public ExcelPreflightCapability[]? Capability { get; set; }

    /// <summary>Include discovered feature rows in object output.</summary>
    [Parameter]
    public SwitchParameter IncludeFeatures { get; set; }

    /// <summary>Include actionable OfficeIMO repair hints for blocked capabilities.</summary>
    [Parameter]
    public SwitchParameter IncludeRepairHints { get; set; }

    /// <summary>Return OfficeIMO's Markdown preflight report instead of structured objects.</summary>
    [Parameter]
    public SwitchParameter AsMarkdown { get; set; }

    /// <summary>Throw when any requested capability is unavailable. Without -Capability, throws when advanced features need review.</summary>
    [Parameter]
    public SwitchParameter ThrowOnFailure { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        using var workbook = ExcelWorkbookCommandService.ResolveWorkbook(this, ParameterSetName, Path, Document, readOnly: true);
        var document = workbook.Document;
        var report = document.InspectFeatures();
        var capabilities = ResolveCapabilities();

        if (ThrowOnFailure.IsPresent) {
            if (Capability != null && Capability.Length > 0) {
                foreach (var capability in capabilities) {
                    report.EnsureCan(capability);
                }
            } else {
                report.EnsureNoAdvancedFeatures();
            }
        }

        if (AsMarkdown.IsPresent) {
            WriteObject(report.ToMarkdown());
            return;
        }

        WriteObject(CreatePreflightObject(document, report, capabilities, IncludeFeatures.IsPresent, IncludeRepairHints.IsPresent));
    }

    private ExcelPreflightCapability[] ResolveCapabilities() {
        if (Capability != null && Capability.Length > 0) {
            return Capability;
        }

        return Enum.GetValues(typeof(ExcelPreflightCapability))
            .Cast<ExcelPreflightCapability>()
            .ToArray();
    }

    private static PSObject CreatePreflightObject(
        ExcelDocument document,
        ExcelFeatureReport report,
        ExcelPreflightCapability[] capabilities,
        bool includeFeatures,
        bool includeRepairHints) {
        var result = new PSObject();
        result.Properties.Add(new PSNoteProperty("Path", document.FilePath));
        result.Properties.Add(new PSNoteProperty("HasAdvancedFeatures", report.HasAdvancedFeatures));
        result.Properties.Add(new PSNoteProperty("FeatureCount", report.Features.Count));
        result.Properties.Add(new PSNoteProperty("EditableFeatureCount", report.EditableFeatures.Count));
        result.Properties.Add(new PSNoteProperty("PartiallyEditableFeatureCount", report.PartiallyEditableFeatures.Count));
        result.Properties.Add(new PSNoteProperty("PreservedFeatureCount", report.PreservedFeatures.Count));
        result.Properties.Add(new PSNoteProperty("UnsupportedFeatureCount", report.UnsupportedFeatures.Count));
        result.Properties.Add(new PSNoteProperty("Capabilities", capabilities.Select(capability => CreateCapabilityObject(report, capability, includeRepairHints)).ToArray()));

        if (includeFeatures) {
            result.Properties.Add(new PSNoteProperty("Features", report.Features.Select(CreateFeatureObject).ToArray()));
        }

        return result;
    }

    private static PSObject CreateCapabilityObject(ExcelFeatureReport report, ExcelPreflightCapability capability, bool includeRepairHints) {
        var diagnostics = report.GetCapabilityDiagnostics(capability).ToArray();
        var item = new PSObject();
        item.Properties.Add(new PSNoteProperty("Name", capability.ToString()));
        item.Properties.Add(new PSNoteProperty("Capability", capability));
        item.Properties.Add(new PSNoteProperty("CanAttempt", report.Can(capability)));
        item.Properties.Add(new PSNoteProperty("Diagnostics", diagnostics));
        item.Properties.Add(new PSNoteProperty("DiagnosticText", diagnostics.Length == 0 ? string.Empty : string.Join("; ", diagnostics)));
        if (includeRepairHints) {
            item.Properties.Add(new PSNoteProperty("RepairHints", report.GetRepairHints(capability).Select(CreateRepairHintObject).ToArray()));
        }

        return item;
    }

    private static PSObject CreateRepairHintObject(ExcelPreflightRepairHint hint) {
        var item = new PSObject();
        item.Properties.Add(new PSNoteProperty("Capability", hint.Capability.ToString()));
        item.Properties.Add(new PSNoteProperty("FeatureName", hint.FeatureName));
        item.Properties.Add(new PSNoteProperty("Action", hint.Action));
        item.Properties.Add(new PSNoteProperty("Command", hint.Command));
        item.Properties.Add(new PSNoteProperty("Details", hint.Details));
        return item;
    }

    private static PSObject CreateFeatureObject(ExcelFeatureFinding feature) {
        var item = new PSObject();
        item.Properties.Add(new PSNoteProperty("Category", feature.Category));
        item.Properties.Add(new PSNoteProperty("Name", feature.Name));
        item.Properties.Add(new PSNoteProperty("Count", feature.Count));
        item.Properties.Add(new PSNoteProperty("SupportLevel", feature.SupportLevel.ToString()));
        item.Properties.Add(new PSNoteProperty("Scope", feature.Scope));
        item.Properties.Add(new PSNoteProperty("Note", feature.Note));
        item.Properties.Add(new PSNoteProperty("Details", feature.Details.ToArray()));
        return item;
    }
}