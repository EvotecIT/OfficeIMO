using System;
using System.Collections.Generic;
using System.IO;

namespace OfficeIMO.Drawing;

/// <summary>
/// Reusable final codec in a raster pipeline: it delegates formats not handled by Drawing to an
/// optional codec, then emits a visible placeholder and diagnostic instead of silently dropping content.
/// </summary>
public sealed class OfficeRasterImageFallbackCodec : IOfficeRasterImageCodec {
    private readonly IOfficeRasterImageCodec? _innerCodec;
    private readonly ICollection<OfficeImageExportDiagnostic>? _diagnostics;
    private readonly string? _source;

    /// <summary>Creates a fallback around an optional application-supplied decoder.</summary>
    public OfficeRasterImageFallbackCodec(
        IOfficeRasterImageCodec? innerCodec = null,
        ICollection<OfficeImageExportDiagnostic>? diagnostics = null,
        string? source = null) {
        _innerCodec = innerCodec;
        _diagnostics = diagnostics;
        _source = source;
    }

    /// <inheritdoc />
    public bool TryDecode(byte[] encodedBytes, string? contentType, out OfficeRasterImage? image) {
        if (_innerCodec != null) {
            try {
                if (_innerCodec.TryDecode((byte[])encodedBytes.Clone(), contentType, out image) && image != null) {
                    AddCallerCodecDiagnostic(contentType);
                    return true;
                }
            } catch (Exception exception) when (
                exception is ArgumentException ||
                exception is FormatException ||
                exception is InvalidOperationException ||
                exception is IOException ||
                exception is NotSupportedException ||
                exception is OverflowException) {
                AddDiagnostic(contentType, exception.Message);
                image = CreatePlaceholder();
                return true;
            }
        }

        AddDiagnostic(contentType, null);
        image = CreatePlaceholder();
        return true;
    }

    private void AddCallerCodecDiagnostic(string? contentType) {
        string format = string.IsNullOrWhiteSpace(contentType) ? "unknown image data" : contentType!;
        _diagnostics?.Add(new OfficeImageExportDiagnostic(
            OfficeImageExportDiagnosticSeverity.Info,
            OfficeImageExportDiagnosticCodes.SourceImageDecodedByCallerCodec,
            "The caller-supplied codec decoded " + format + ".",
            _source));
    }

    private void AddDiagnostic(string? contentType, string? detail) {
        string format = string.IsNullOrWhiteSpace(contentType) ? "unknown image data" : contentType!;
        string message = "Drawing could not decode " + format + "; a visible placeholder was rendered.";
        if (!string.IsNullOrWhiteSpace(detail)) message += " " + detail;
        _diagnostics?.Add(new OfficeImageExportDiagnostic(
            OfficeImageExportDiagnosticSeverity.Warning,
            OfficeImageExportDiagnosticCodes.SourceImageDecodeFallback,
            message,
            _source));
    }

    private static OfficeRasterImage CreatePlaceholder() {
        const int size = 32;
        var image = new OfficeRasterImage(size, size, OfficeColor.FromRgb(245, 245, 245));
        OfficeColor border = OfficeColor.FromRgb(150, 150, 150);
        OfficeColor cross = OfficeColor.FromRgb(190, 70, 70);
        for (int offset = 0; offset < size; offset++) {
            image.SetPixel(offset, 0, border);
            image.SetPixel(offset, size - 1, border);
            image.SetPixel(0, offset, border);
            image.SetPixel(size - 1, offset, border);
            image.SetPixel(offset, offset, cross);
            image.SetPixel(size - 1 - offset, offset, cross);
        }
        return image;
    }
}
