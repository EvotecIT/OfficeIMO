using System.Management.Automation;
using System.Text;
using OfficeIMO.Core.Internal;
using OfficeIMO.Pdf;
using PSWriteOffice.Foundation;
using PSWriteOffice.Services.Pdf;

namespace PSWriteOffice.Cmdlets.Pdf;

/// <summary>Exports native PDF form values as XFDF XML or atomically published XFDF files.</summary>
/// <example><summary>Export each form in a folder.</summary><code>Export-OfficePdfXfdf -Path './Forms/*.pdf' -OutputDirectory ./Values -NoClobber -PassThru</code></example>
[Cmdlet(VerbsData.Export, "OfficePdfXfdf", DefaultParameterSetName = "Document", SupportsShouldProcess = true)]
[OutputType(typeof(string), typeof(FileInfo))]
public sealed partial class ExportOfficePdfXfdfCommand : PSCmdlet {
    private readonly List<string> _inputs = new();
    /// <inheritdoc />
    protected override void ProcessRecord() {
        if (ParameterSetName == "Document") {
            var document = Document ?? OfficeDocumentContext.Find<PdfDocument>() ?? throw new PSArgumentException("Provide Document or use an Open-/New-OfficePdf block.");
            WriteObject(document.Forms.ExportXfdf());
            return;
        }
        bool literal = ParameterSetName.StartsWith("LiteralPath", StringComparison.Ordinal);
        _inputs.AddRange(OfficePath.ResolveInputs(this, literal ? LiteralPath : Path, literal));
    }
    /// <inheritdoc />
    protected override void EndProcessing() {
        if (_inputs.Count == 0) return;
        if (OutputPath != null && _inputs.Count != 1) throw new PSArgumentException("OutputPath requires one PDF; use OutputDirectory for a batch.");
        var read = PdfCommandUtilities.CreateReadOptions(LoadOptions, Password, IgnorePermissionRestrictions.IsPresent);
        if (OfficeSwitch.Get(MyInvocation.BoundParameters, nameof(IgnorePermissionRestrictions)) is bool ignore)
            read = (read ?? new PdfLoadOptions()).Snapshot(null, null, ignore ? PdfPermissionPolicy.IgnoreRestrictions : PdfPermissionPolicy.Enforce);
        var destinations = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (string input in _inputs) {
            string? output = OutputPath == null ? OutputDirectory == null ? null : System.IO.Path.Combine(
                OfficePath.ResolveOutput(SessionState, OutputDirectory, true), System.IO.Path.GetFileNameWithoutExtension(input) + ".xfdf")
                : OfficePath.ResolveOutput(SessionState, OutputPath, false);
            if (output != null) {
                if (OfficePath.AreEquivalent(input, output)) throw new PSArgumentException("PDF and XFDF paths must differ.");
                if (!destinations.Add(output)) throw new PSArgumentException("Two PDF inputs have the same XFDF destination.");
                OfficePath.ValidateDestination(output, NoClobber.IsPresent);
                if (!ShouldProcess(output, "Export XFDF form data")) continue;
            }
            using var document = PdfDocument.Load(input, read);
            string xml = document.Forms.ExportXfdf();
            if (output == null) { WriteObject(xml); continue; }
            PdfCommandUtilities.EnsureDirectory(output);
            OfficeFileOutput.Write(output, !NoClobber.IsPresent, temporary => File.WriteAllText(temporary, xml, new UTF8Encoding(false)));
            if (PassThru) WriteObject(new FileInfo(output));
        }
    }
}
