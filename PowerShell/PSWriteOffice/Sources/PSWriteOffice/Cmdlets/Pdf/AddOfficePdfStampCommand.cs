using System.Management.Automation;
using System.Security;
using OfficeIMO;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using PSWriteOffice.Foundation;
using PSWriteOffice.Services.Pdf;

namespace PSWriteOffice.Cmdlets.Pdf;

/// <summary>Stamps text, images, native canvas content or a PDF page onto existing PDFs.</summary>
[Cmdlet(VerbsCommon.Add, "OfficePdfStamp", DefaultParameterSetName = "TextDocument", SupportsShouldProcess = true)]
[OutputType(typeof(PdfDocument), typeof(FileInfo))]
public sealed partial class AddOfficePdfStampCommand : PSCmdlet, IDisposable {
    private readonly CancellationTokenSource _stop = new();
    private readonly List<string> _files = new();
    private readonly List<PdfDocument> _documents = new();
    /// <inheritdoc />
    protected override void ProcessRecord() {
        if (ParameterSetName.EndsWith("Document", StringComparison.Ordinal)) {
            _documents.Add(Document ?? OfficeDocumentContext.Find<PdfDocument>() ??
                throw new PSArgumentException("Supply Document or use an enclosing Open-OfficePdf block."));
        } else {
            bool literal = ParameterSetName.Contains("LiteralPath");
            _files.AddRange(OfficePath.ResolveInputs(this, literal ? LiteralPath : Path, literal));
        }
    }
    /// <inheritdoc />
    protected override void EndProcessing() {
        string mode = ParameterSetName.StartsWith("Text", StringComparison.Ordinal) ? "Text" :
            ParameterSetName.StartsWith("Image", StringComparison.Ordinal) ? "Image" :
            ParameterSetName.StartsWith("Content", StringComparison.Ordinal) ? "Content" : "Page";
        foreach (var document in _documents) {
            _stop.Token.ThrowIfCancellationRequested();
            if (!ShouldProcess("PDF document", "Add " + mode.ToLowerInvariant() + " stamp")) continue;
            var result = Stamp(document, mode, null);
            bool adopted = OfficeDocumentContext.Replace(document, result);
            if (PassThru.IsPresent || !adopted) WriteObject(result);
        }
        if (_files.Count == 0) return;
        var work = PdfFileCommandInputs.Resolve(this, null, _files.ToArray(), OutputPath, OutputDirectory, noClobber: false);
        var read = ReadPolicy(LoadOptions, Password, nameof(IgnorePermissionRestrictions));
        foreach (var item in work) {
            _stop.Token.ThrowIfCancellationRequested();
            if (!ShouldProcess(item.Output, "Save PDF with " + mode.ToLowerInvariant() + " stamp")) continue;
            using var source = PdfDocument.Load(item.Input, read, _stop.Token);
            var result = Stamp(source, mode, read);
            try {
                result.Save(item.Output, new PdfSaveOptions { ConflictPolicy = OfficeConversionFileConflictPolicy.Replace }, _stop.Token);
                if (PassThru.IsPresent) WriteObject(new FileInfo(item.Output));
            } finally { if (!ReferenceEquals(source, result)) result.Dispose(); }
        }
    }
    private PdfDocument Stamp(PdfDocument document, string mode, PdfLoadOptions? read) {
        if (mode == "Text") {
            var options = new PdfTextStampOptions { X = X, Y = Y, RotationDegrees = Rotation,
                BehindContent = Behind.IsPresent, Opacity = Opacity };
            if (Size.HasValue) options.FontSize = Size.Value;
            if (Color.HasValue) options.Color = Color.Value;
            if (Page != null) options.PageNumbers = Page.Resolve(document.GetPageLayouts(read, _stop.Token).Count).ToArray();
            return Behind.IsPresent ? document.Stamp.TextWatermark(Text!, options, read, _stop.Token)
                : document.Stamp.Text(Text!, options, read, _stop.Token);
        }
        if (mode == "Image") {
            var options = new PdfImageStampOptions { X = X, Y = Y, Width = Width, Height = Height,
                RotationDegrees = Rotation, BehindContent = Behind.IsPresent, Opacity = Opacity };
            if (Page != null) options.PageNumbers = Page.Resolve(document.GetPageLayouts(read, _stop.Token).Count).ToArray();
            string image = ResolveOne(ImagePath!);
            using var stream = File.OpenRead(image);
            return Behind.IsPresent ? document.Stamp.ImageWatermark(stream, options, read, _stop.Token)
                : document.Stamp.Image(stream, options, read, _stop.Token);
        }
        if (mode == "Content") {
            PdfCreateOptions? rendering = null;
            if (ConfigureRendering != null) { rendering = new PdfCreateOptions(); ConfigureRendering.Invoke(rendering); }
            var options = new PdfCanvasStampOptions { BehindContent = Behind.IsPresent, Opacity = Opacity,
                TargetPages = Page, RenderingOptions = rendering };
            var transform = OfficeTransform.RotateDegrees(Rotation).Then(OfficeTransform.Translate(X?.ToPoints() ?? 0D, Y?.ToPoints() ?? 0D));
            return document.Stamp.Content((canvas, page) => canvas.Effect(transform, 1D, nested => {
                _stop.Token.ThrowIfCancellationRequested();
                using (PdfCanvasDslContext.Enter(nested, page)) Content!.Invoke(nested, page);
            }), options, read, _stop.Token);
        }
        var overlayRead = ReadPolicy(OverlayLoadOptions, OverlayPassword, nameof(IgnoreOverlayPermissionRestrictions));
        var overlay = new PdfPageOverlayOptions { SourcePageNumber = OverlayPage, Fit = Fit,
            HorizontalAlignment = HorizontalAlignment,
            VerticalAlignment = VerticalAlignment == OfficeTextVerticalAlignment.Top ? PdfVerticalAlign.Top :
                VerticalAlignment == OfficeTextVerticalAlignment.Bottom ? PdfVerticalAlign.Bottom : PdfVerticalAlign.Middle,
            X = X?.ToPoints(), Y = Y?.ToPoints(), Width = Width?.ToPoints(), Height = Height?.ToPoints(),
            Opacity = Opacity, RotationDegrees = Rotation, BehindContent = Behind.IsPresent, SourceReadOptions = overlayRead,
            TargetPages = Page };
        string path = ResolveOne(OverlayPath!);
        return Behind.IsPresent ? document.Stamp.UnderlayPage(path, overlay, read, _stop.Token)
            : document.Stamp.OverlayPage(path, overlay, read, _stop.Token);
    }
    private string ResolveOne(string value) => OfficePath.ResolveInput(SessionState, value, literal: true)[0];
    private PdfLoadOptions? ReadPolicy(PdfLoadOptions? options, SecureString? password, string ignoreName) {
        var result = PdfCommandUtilities.CreateReadOptions(options, password);
        if (OfficeSwitch.Get(MyInvocation.BoundParameters, ignoreName) is bool ignore)
            result = (result ?? new PdfLoadOptions()).Snapshot(null, null, ignore ? PdfPermissionPolicy.IgnoreRestrictions : PdfPermissionPolicy.Enforce);
        return result;
    }
    /// <inheritdoc />
    protected override void StopProcessing() => _stop.Cancel();
    /// <summary>Releases this invocation's cancellation source.</summary>
    public void Dispose() => _stop.Dispose();
}
