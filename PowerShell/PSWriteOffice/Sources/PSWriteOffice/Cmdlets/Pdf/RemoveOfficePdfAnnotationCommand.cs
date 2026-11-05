using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Management.Automation;
using System.Threading;
using OfficeIMO;
using OfficeIMO.Pdf;
using PSWriteOffice.Foundation;
using PSWriteOffice.Services.Pdf;

namespace PSWriteOffice.Cmdlets.Pdf;

/// <summary>Removes native annotations selected by identity, page and subtype.</summary>
/// <para>Native confirmation runs before loading or editing a file. Live inputs stay caller-owned; a replacement document is adopted by the enclosing content block or returned for explicit assignment. PassThru returns native edit reports.</para>
[Cmdlet(VerbsCommon.Remove, "OfficePdfAnnotation", DefaultParameterSetName = "Document", SupportsShouldProcess = true)]
[OutputType(typeof(FileInfo), typeof(PdfDocument), typeof(PdfAnnotationEditResult))]
public sealed partial class RemoveOfficePdfAnnotationCommand : PSCmdlet, IDisposable {
    private readonly CancellationTokenSource _stop = new();
    private readonly List<string> _inputs = new();

    /// <inheritdoc />
    protected override void ProcessRecord() {
        if (ParameterSetName.StartsWith("Document", StringComparison.Ordinal)) {
            var source = Document ?? OfficeDocumentContext.Find<PdfDocument>() ?? throw new PSArgumentException("Provide Document or use an Open-/New-OfficePdf content block.");
            if (!ShouldProcess("PDF document", "Remove PDF annotations")) return;
            var reports = Edit(source);
            if (reports.Count == 0) { if (!PassThru) WriteObject(source); return; }
            var result = reports[reports.Count - 1].ToDocument();
            bool adopted = OfficeDocumentContext.Replace(source, result);
            if (PassThru) {
                foreach (var report in reports) WriteObject(report);
                if (!adopted) result.Dispose();
            } else if (!adopted) WriteObject(result);
            return;
        }
        bool literal = ParameterSetName.StartsWith("LiteralPath", StringComparison.Ordinal);
        _inputs.AddRange(OfficePath.ResolveInputs(this, literal ? LiteralPath : Path, literal));
    }

    /// <inheritdoc />
    protected override void EndProcessing() {
        if (_inputs.Count == 0) return;
        var read = PdfCommandUtilities.CreateReadOptions(LoadOptions, Password, IgnorePermissionRestrictions.IsPresent);
        if (OfficeSwitch.Get(MyInvocation.BoundParameters, nameof(IgnorePermissionRestrictions)) is bool ignore)
            read = (read ?? new PdfLoadOptions()).Snapshot(null, null, ignore ? PdfPermissionPolicy.IgnoreRestrictions : PdfPermissionPolicy.Enforce);
        foreach (var item in PdfFileCommandInputs.Resolve(this, null, _inputs.ToArray(), OutputPath, OutputDirectory, false)) {
            if (!ShouldProcess(item.Output, "Remove PDF annotations")) continue;
            using var source = PdfDocument.Load(item.Input, read, _stop.Token);
            var reports = Edit(source);
            using var result = reports.Count == 0 ? null : reports[reports.Count - 1].ToDocument();
            (result ?? source).Save(item.Output, new PdfSaveOptions { ConflictPolicy = OfficeConversionFileConflictPolicy.Replace }, _stop.Token);
            if (PassThru) { foreach (var report in reports) WriteObject(report); }
            else WriteObject(new FileInfo(item.Output));
        }
    }

    /// <inheritdoc />
    protected override void StopProcessing() => _stop.Cancel();
    private IReadOnlyList<PdfAnnotationEditResult> Edit(PdfDocument source) {
        var options = new PdfAnnotationRemovalOptions { ObjectNumber = ObjectNumber, PageNumber = Page,
            Subtype = Subtype, RemoveMatchingPopups = !KeepPopups.IsPresent };
        return new[] { source.Annotations.Remove(options, _stop.Token) };
    }
    /// <summary>Releases the command-owned cancellation source.</summary>
    public void Dispose() => _stop.Dispose();
}
