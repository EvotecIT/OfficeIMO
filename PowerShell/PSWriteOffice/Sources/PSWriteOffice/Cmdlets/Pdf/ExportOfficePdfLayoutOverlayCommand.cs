using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Management.Automation;
using OfficeIMO.Core.Internal;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using PSWriteOffice.Foundation;
using PSWriteOffice.Services.Pdf;

namespace PSWriteOffice.Cmdlets.Pdf;

/// <summary>Exports native PDF word, line, region and reading-order diagnostics as PNG or SVG.</summary>
/// <example><summary>Export literal-file diagnostics.</summary><code>Export-OfficePdfLayoutOverlay -LiteralPath './Report[1].pdf' -OutputPath ./Report-layout.svg -Page 1 -Format Svg -PassThru</code></example>
[Cmdlet(VerbsData.Export, "OfficePdfLayoutOverlay", DefaultParameterSetName = "Path", SupportsShouldProcess = true)]
[OutputType(typeof(OfficeImageExportResult))]
public sealed partial class ExportOfficePdfLayoutOverlayCommand : OfficeCmdlet {
    private readonly List<string> _inputs = new();
    /// <inheritdoc />
    protected override void ProcessRecord() {
        bool literal = ParameterSetName.StartsWith("LiteralPath", StringComparison.Ordinal);
        _inputs.AddRange(OfficePath.ResolveInputs(this, literal ? LiteralPath : Path, literal));
    }
    /// <inheritdoc />
    protected override void EndProcessing() {
        if (Format is not (OfficeImageExportFormat.Png or OfficeImageExportFormat.Svg)) throw new PSArgumentException("Layout overlays support PNG and SVG.");
        if (Page < 1 || Scale <= 0 || double.IsNaN(Scale) || double.IsInfinity(Scale)) throw new PSArgumentException("Page and Scale must be positive finite values.");
        var binder = new OfficeOptionsTransformationAttribute(typeof(PdfLayoutDebugOverlayOptions));
        var overlay = new PdfLayoutDebugOverlayOptions();
        foreach (object item in Options ?? Array.Empty<object>()) {
            object? value = item is PSObject wrapped ? wrapped.BaseObject : item;
            if (value is PdfLayoutDebugOverlayOptions native) overlay = native.Clone();
            else if (value is IDictionary fields) overlay = (PdfLayoutDebugOverlayOptions)binder.ApplyTo(
                (EngineIntrinsics)SessionState.PSVariable.GetValue("ExecutionContext"), overlay, fields);
            else throw new PSArgumentException("Options accepts PdfLayoutDebugOverlayOptions or its property hashtables.");
        }
        var read = PdfCommandUtilities.CreateReadOptions(LoadOptions, Password, IgnorePermissionRestrictions.IsPresent);
        if (OfficeSwitch.Get(MyInvocation.BoundParameters, nameof(IgnorePermissionRestrictions)) is bool ignore)
            read = (read ?? new PdfLoadOptions()).Snapshot(null, null, ignore ? PdfPermissionPolicy.IgnoreRestrictions : PdfPermissionPolicy.Enforce);
        foreach (var item in PdfFileCommandInputs.Resolve(this, null, _inputs.ToArray(), OutputPath, OutputDirectory, NoClobber, Format == OfficeImageExportFormat.Svg ? ".svg" : ".png")) {
            if (OfficePath.AreEquivalent(item.Input, item.Output)) throw new PSArgumentException("PDF and overlay paths must differ.");
            if (!ShouldProcess(item.Output, "Export PDF layout overlay")) continue;
            using var document = PdfDocument.Load(item.Input, read);
            var drawing = Format == OfficeImageExportFormat.Svg ? document.Render.LayoutDebugOverlay(Page, overlay, LayoutOptions, read) : null;
            var bytes = Format == OfficeImageExportFormat.Svg ? OfficeDrawingSvgExporter.ToSvgBytes(drawing!, Scale, OfficeSvgSizeUnit.Pixel)
                : document.Render.LayoutDebugOverlayPng(Page, Scale, overlay, LayoutOptions, read);
            var size = Format == OfficeImageExportFormat.Png ? OfficeImageReader.Identify(bytes, "overlay.png") : null;
            var result = new OfficeImageExportResult(Format, size?.Width ?? checked((int)Math.Ceiling(drawing!.Width * Scale)),
                size?.Height ?? checked((int)Math.Ceiling(drawing!.Height * Scale)), bytes, $"Page {Page} layout", $"{item.Input}#page={Page}");
            PdfCommandUtilities.EnsureDirectory(item.Output);
            OfficeFileOutput.Write(item.Output, !NoClobber.IsPresent, temporary => File.WriteAllBytes(temporary, bytes));
            if (PassThru) WriteObject(result);
        }
    }
}
