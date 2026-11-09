using System.Management.Automation;
using OfficeIMO;
using OfficeIMO.Pdf;
using PSWriteOffice.Foundation;
using PSWriteOffice.Services.Pdf;

namespace PSWriteOffice.Cmdlets.Pdf;

/// <summary>Updates native annotation contents, styling and actions by identity or target URI.</summary>
/// <para>Native confirmation runs before loading or editing a file. Live inputs stay caller-owned; a replacement document is adopted by the enclosing content block or returned for explicit assignment. PassThru returns native edit reports.</para>
[Cmdlet(VerbsCommon.Set, "OfficePdfAnnotation", DefaultParameterSetName = "Document", SupportsShouldProcess = true)]
[OutputType(typeof(FileInfo), typeof(PdfDocument), typeof(PdfAnnotationEditResult))]
public sealed partial class SetOfficePdfAnnotationCommand : PSCmdlet, IDisposable {
    private readonly CancellationTokenSource _stop = new();
    private readonly List<string> _inputs = new();

    /// <inheritdoc />
    protected override void ProcessRecord() {
        if (ParameterSetName.StartsWith("Document", StringComparison.Ordinal)) {
            var source = Document ?? OfficeDocumentContext.Find<PdfDocument>() ?? throw new PSArgumentException("Provide Document or use an Open-/New-OfficePdf content block.");
            if (!ShouldProcess("PDF document", "Set PDF annotations")) return;
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
            if (!ShouldProcess(item.Output, "Set PDF annotations")) continue;
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
        var options = PdfOperationCancellation.Bind(Options ?? new PdfAnnotationUpdateOptions(), createWhenNull: false, _stop.Token)!;
        if (MyInvocation.BoundParameters.ContainsKey(nameof(Contents))) options.Contents = Contents;
        if (MyInvocation.BoundParameters.ContainsKey(nameof(Title))) options.Title = Title;
        if (MyInvocation.BoundParameters.ContainsKey(nameof(Name))) options.Name = Name;
        if (MyInvocation.BoundParameters.ContainsKey(nameof(Flags))) options.Flags = Flags;
        if (MyInvocation.BoundParameters.ContainsKey(nameof(RemoveAction))) options.RemoveActions = RemoveAction.IsPresent;
        if (Color.HasValue) {
            var rgb = Color.Value.Resolve(OfficeTheme.Default).ToNormalizedRgb();
            options.Color = new[] { rgb.R, rgb.G, rgb.B };
        }
        if (Uri == null) return new[] { source.Annotations.Update(ObjectNumber, options, _stop.Token) };
        var pattern = new WildcardPattern(Uri, WildcardOptions.IgnoreCase | WildcardOptions.CultureInvariant);
        var selected = source.Annotations.GetForEditing(_stop.Token)
            .Where(annotation => annotation.Uri != null && pattern.IsMatch(annotation.Uri) && annotation.ObjectNumber.HasValue)
            .Select(annotation => annotation.ObjectNumber!.Value).ToList();
        var reports = new List<PdfAnnotationEditResult>();
        PdfDocument current = source;
        try {
            while (selected.Count > 0) {
                _stop.Token.ThrowIfCancellationRequested();
                int number = selected[0]; selected.RemoveAt(0);
                var report = current.Annotations.Update(number, options, _stop.Token);
                reports.Add(report);
                var next = report.ToDocument();
                if (!ReferenceEquals(current, source)) current.Dispose();
                current = next;
                if (report.AnnotationObjectNumberMap != null)
                    selected = selected.Select(id => report.AnnotationObjectNumberMap.TryGetValue(id, out int mapped) ? mapped
                        : throw new InvalidOperationException("An annotation disappeared during the native rewrite.")).ToList();
            }
        } finally { if (!ReferenceEquals(current, source)) current.Dispose(); }
        return reports;
    }
    /// <summary>Releases the command-owned cancellation source.</summary>
    public void Dispose() => _stop.Dispose();
}
