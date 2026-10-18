using System.IO;
using System.Management.Automation;
using OfficeIMO.Excel;
using PSWriteOffice.Services.Excel;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Sets or resets the workbook theme package part for an Excel workbook.</summary>
/// <example>
///   <summary>Reset and rename the workbook theme.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$theme = Set-OfficeExcelTheme -Path .\Report.xlsx -Default -Name 'Contoso Workbook Theme' -PassThru
/// Get-OfficeExcelSummary -Path .\Report.xlsx |
///     Select-Object Path, WorksheetCount</code>
///   <para>Writes the built-in OfficeIMO workbook theme and updates its theme name.</para>
/// </example>
[Cmdlet(VerbsCommon.Set, "OfficeExcelTheme", DefaultParameterSetName = ParameterSetContext, SupportsShouldProcess = true)]
[Alias("ExcelTheme")]
[OutputType(typeof(PSObject))]
public sealed class SetOfficeExcelThemeCommand : PSCmdlet {
    private const string ParameterSetContext = "Context";
    private const string ParameterSetDocument = "Document";
    private const string ParameterSetPath = "Path";

    /// <summary>Workbook path to update.</summary>
    [Parameter(Mandatory = true, Position = 0, ParameterSetName = ParameterSetPath)]
    [Alias("InputPath", "FilePath")]
    public string Path { get; set; } = string.Empty;

    /// <summary>Workbook to update outside the DSL context.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = ParameterSetDocument)]
    public ExcelDocument Document { get; set; } = null!;

    /// <summary>Reset the workbook to the built-in OfficeIMO theme.</summary>
    [Parameter]
    public SwitchParameter Default { get; set; }

    /// <summary>Theme XML to write to the workbook theme part.</summary>
    [Parameter]
    public string? Xml { get; set; }

    /// <summary>Path to a DrawingML theme XML file.</summary>
    [Parameter]
    public string? XmlPath { get; set; }

    /// <summary>Optional workbook theme name to apply after writing the theme.</summary>
    [Parameter]
    public string? Name { get; set; }

    /// <summary>Emit workbook theme metadata after applying the update.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        using var workbook = ExcelWorkbookCommandService.ResolveWorkbook(this, ParameterSetName, Path, Document, readOnly: false);
        if (!ExcelShouldProcessService.ShouldProcessWorkbook(this, workbook.Document, Path, "Update Excel workbook")) {
            return;
        }


        string? xml = ResolveThemeXml();
        if (Default.IsPresent) {
            workbook.Document.ResetWorkbookTheme(Name);
        } else if (xml != null) {
            workbook.Document.SetWorkbookThemeXml(xml);
            if (!string.IsNullOrWhiteSpace(Name)) {
                workbook.Document.SetWorkbookThemeName(Name!);
            }
        } else if (!string.IsNullOrWhiteSpace(Name)) {
            workbook.Document.SetWorkbookThemeName(Name!);
        } else {
            throw new PSArgumentException("Specify -Default, -Xml, -XmlPath, or -Name.");
        }

        ExcelWorkbookThemeInfo info = workbook.Document.GetWorkbookTheme(includeXml: false);
        workbook.SaveIfOwned();

        if (PassThru.IsPresent) {
            var result = new PSObject();
            result.Properties.Add(new PSNoteProperty("HasTheme", info.HasTheme));
            result.Properties.Add(new PSNoteProperty("Name", info.Name));
            WriteObject(result);
        }
    }

    private string? ResolveThemeXml() {
        if (!string.IsNullOrWhiteSpace(Xml) && !string.IsNullOrWhiteSpace(XmlPath)) {
            throw new PSArgumentException("Specify either Xml or XmlPath, not both.");
        }

        if (!string.IsNullOrWhiteSpace(Xml)) {
            return Xml;
        }

        if (string.IsNullOrWhiteSpace(XmlPath)) {
            return null;
        }

        string resolved = SessionState.Path.GetUnresolvedProviderPathFromPSPath(XmlPath!);
        return File.ReadAllText(resolved);
    }
}