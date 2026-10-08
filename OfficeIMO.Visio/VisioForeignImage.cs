using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Xml.Linq;
using OfficeIMO.Drawing;

namespace OfficeIMO.Visio;

// Foreign images are page content, separate from decorative stencil preview artwork.
internal static partial class VisioForeignImage {
    private static readonly XNamespace Ns = "http://schemas.microsoft.com/office/visio/2012/main";
    private static readonly XNamespace Relationships = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    internal static bool IsForeign(VisioShape shape) => string.Equals(shape.Type ?? MasterShape(shape)?.Type, "Foreign", StringComparison.OrdinalIgnoreCase);
    private static VisioShape? MasterShape(VisioShape shape) => shape.MasterShape ?? shape.Master?.Shape;
    private static XElement? ForeignData(VisioShape shape) => shape.PreservedShapeChildren
        .Select(entry => entry.RawElement).FirstOrDefault(element => element?.Name == Ns + "ForeignData");

    internal static OfficeRasterImage Decode(VisioShape shape, IOfficeRasterImageCodec? codec,
        ICollection<OfficeImageExportDiagnostic>? diagnostics, string? source, CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        VisioShape owner = shape;
        XElement? foreign = ForeignData(owner);
        if (foreign == null && MasterShape(shape) is VisioShape master) { owner = master; foreign = ForeignData(master); }
        string? id = (string?)foreign?.Element(Ns + "Rel")?.Attribute(Relationships + "id");
        VisioForeignResource? resource = owner.ForeignResources.FirstOrDefault(candidate => candidate.RelationshipId == id);
        bool image = resource != null && resource.RelationshipType.EndsWith("/image", StringComparison.Ordinal) &&
            (string?)foreign?.Attribute("ForeignType") is "Bitmap" or "Metafile" or "EnhMetaFile";
        string location = Location(shape, source);
        if (image) {
            byte[] bytes = PrepareBitmap(resource!.Bytes);
            if (!ReferenceEquals(bytes, resource.Bytes))
                diagnostics?.Add(new OfficeImageExportDiagnostic(OfficeImageExportDiagnosticSeverity.Info,
                    "VISIO_BITMAP_SIZE_NORMALIZED", "A zero BMP file-size field was supplied for decoding; the original embedded bytes remain unchanged.", location));
            var options = new OfficeRasterDecodeOptions {
                MaximumEncodedBytes = 64 * 1024 * 1024, MaximumDecodedPixels = 32_000_000,
                CancellationToken = cancellationToken, ImageCodec = codec
            };
            if (OfficeRasterImageDecoder.TryDecode(bytes, options, out OfficeRasterImage? raster, out OfficeRasterDecodeInfo info) && raster != null) {
                if (info.UsedCallerCodec)
                    new OfficeRasterImageFallbackCodec(null, diagnostics, location).AddCallerCodecDiagnostic(resource.ContentType);
                if (info.FramesOrPagesDiscarded || info.AnimationDiscarded)
                    diagnostics?.Add(new OfficeImageExportDiagnostic(OfficeImageExportDiagnosticSeverity.Warning,
                        OfficeImageExportDiagnosticCodes.SourceImageStaticFrameSelected,
                        info.Diagnostic ?? "Only the first image frame or page is rendered.", location, OfficeConversionLossKind.Omission));
                return raster;
            }
            // Metafiles use the explicit vector-codec boundary after native record validation.
            // Rejected raster inputs must not bypass the shared inspector through this fallback.
            if (bytes.Length <= options.MaximumEncodedBytes &&
                OfficeImageReader.TryIdentifyByContent(bytes, null, out OfficeImageInfo vector) &&
                vector.Format is OfficeImageFormat.Wmf or OfficeImageFormat.Emf) {
                cancellationToken.ThrowIfCancellationRequested();
                var fallback = new OfficeRasterImageFallbackCodec(codec, diagnostics, location);
                fallback.TryDecode(bytes, resource.ContentType, out raster);
                cancellationToken.ThrowIfCancellationRequested();
                if (raster != null && (long)raster.Width * raster.Height <= options.MaximumDecodedPixels) return raster;
            }
        }
        // OLE and unrecognized payloads never enter the image codec, even when they resemble images.
        new OfficeRasterImageFallbackCodec(null, diagnostics, location).TryDecode(Array.Empty<byte>(),
            image ? "unsupported image or image outside the decoding limits" : "unsupported or unavailable Visio foreign content", out OfficeRasterImage? placeholder);
        return placeholder!;
    }

    private static string Location(VisioShape shape, string? source) =>
        (source ?? "Visio page") + " / " + (shape.NameU ?? shape.Id);

    // Some independent VDX producers embed BMP streams with an unspecified bfSize.
    // Supply the bounded payload length only for decoding; all other BMP checks remain shared.
    internal static byte[] PrepareBitmap(byte[] bytes) {
        if (bytes.Length < 54 || bytes[0] != 'B' || bytes[1] != 'M' || bytes[2] != 0 || bytes[3] != 0 || bytes[4] != 0 || bytes[5] != 0) return bytes;
        byte[] copy = (byte[])bytes.Clone();
        for (int index = 0; index < 4; index++) copy[2 + index] = (byte)(bytes.Length >> (index * 8));
        return copy;
    }
}
