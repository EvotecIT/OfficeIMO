using System;
using System.IO;
using System.Management.Automation;
using System.Text;
using OfficeIMO.Core.Internal;
using OfficeIMO.Project;
using PSWriteOffice.Services;
#if !FRAMEWORK
using OfficeIMO.Workflows;
#endif

namespace PSWriteOffice.Cmdlets.Project;

/// <summary>Exports a captured Project report through the shared PDF, HTML, Word, Excel or PowerPoint adapters.</summary>
/// <para>Requires PowerShell 7.4. The destination extension selects the report adapter. Rendering and editable document composition remain in OfficeIMO.Workflows.</para>
/// <example>
///   <summary>Create an editable report and a vector PDF from the same schedule view.</summary>
///   <code>$view = $project | Get-OfficeProjectView -Kind Gantt
/// $view | Export-OfficeProjectReport -Path ./delivery.pdf
/// $view | Export-OfficeProjectReport -Path ./delivery.docx</code>
/// </example>
[Cmdlet(VerbsData.Export, "OfficeProjectReport", SupportsShouldProcess = true, DefaultParameterSetName = "View")]
[OutputType(typeof(FileInfo))]
public sealed class ExportOfficeProjectReportCommand : PSCmdlet {
    /// <summary>Captured report from Get-OfficeProjectView.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true)]
    public ProjectView View { get; set; } = null!;

    /// <summary>PDF, HTML, DOCX, XLSX or PPTX destination filename.</summary>
    [Parameter(Mandatory = true, Position = 0)]
    public string Path { get; set; } = string.Empty;

    /// <summary>Replace an existing destination.</summary>
    [Parameter]
    public SwitchParameter Force { get; set; }

    /// <summary>Return the saved report file.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
#if FRAMEWORK
        throw new PlatformNotSupportedException("Project report export requires PowerShell 7.4 or newer.");
#else
        string path = OfficeFileCommandUtilities.ResolvePath(this, Path);
        if (!ShouldProcess(path, "Export project report")) return;
        byte[] bytes;
        switch (System.IO.Path.GetExtension(path).ToLowerInvariant()) {
            case ".pdf": bytes = ProjectReportWorkflow.ToPdf(View); break;
            case ".html": case ".htm": bytes = new UTF8Encoding(false).GetBytes(ProjectReportWorkflow.ToHtml(View)); break;
            case ".docx":
                using (var document = ProjectReportWorkflow.CreateWord(View)) bytes = document.ToBytes();
                break;
            case ".xlsx":
                using (var document = ProjectReportWorkflow.CreateExcel(View)) bytes = document.ToBytes();
                break;
            case ".pptx":
                using (var document = ProjectReportWorkflow.CreatePowerPoint(View)) bytes = document.ToBytes();
                break;
            default: throw new PSArgumentException("Choose a .pdf, .html, .docx, .xlsx or .pptx report destination.", nameof(Path));
        }
        OfficeFileOutput.Write(path, bytes, Force.IsPresent);
        if (PassThru.IsPresent) WriteObject(new FileInfo(path));
#endif
    }
}
