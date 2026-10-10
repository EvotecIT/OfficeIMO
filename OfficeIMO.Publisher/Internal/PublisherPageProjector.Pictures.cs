using OfficeIMO.Drawing;
using OfficeIMO.Drawing.Binary;

namespace OfficeIMO.Publisher.Internal;

internal sealed partial class PublisherPageProjector {
    private OfficeArtPictureEffectProjector PictureEffects(OfficeArtPictureProperties picture, uint objectId) {
        OfficeArtPictureEffectProjector effects = OfficeArtPictureEffectProjector.Create(picture,
            reference => reference.TryResolve(index => index < _source.Palette.Count ? _source.Palette[index] : null,
                out OfficeColor color) ? color : null);
        string location = PublisherEscherReader.ShapeLocation(objectId);
        if ((effects.Limits & (OfficeArtPictureEffectLimit.InvalidBrightness | OfficeArtPictureEffectLimit.InvalidContrast)) != 0)
            _context.Add("PUB_PICTURE_TONE_INVALID", "An invalid native brightness or contrast control was omitted; other valid picture controls remain applicable.",
                OfficeConversionLossKind.Omission, location);
        if ((effects.Limits & OfficeArtPictureEffectLimit.UnresolvedTransparentColor) != 0)
            _context.Add("PUB_PICTURE_TRANSPARENT_COLOR_UNRESOLVED", "The native transparent-color key could not be resolved and was omitted.",
                OfficeConversionLossKind.Omission, location);
        if ((effects.Limits & (OfficeArtPictureEffectLimit.UnqualifiedRecolor | OfficeArtPictureEffectLimit.ExtendedColor)) != 0)
            _context.Add("PUB_PICTURE_RECOLOR_UNASSESSED", "Native recoloring and extended color modifications require additional qualified projection.",
                OfficeConversionLossKind.Unassessed, location);
        return effects;
    }

    private byte[] ProjectImage(PublisherImage image, uint objectId, OfficeArtPictureEffectProjector effects, out string contentType) {
        contentType = image.ContentType;
        OfficeImageFormat format = OfficeImageInfo.FromMimeType(contentType);
        bool metafile = format is OfficeImageFormat.Wmf or OfficeImageFormat.Emf;
        if (!effects.HasProjection && !metafile) return image.Bytes;
        long maximumPixels = _context.Options.MaximumRasterPixels;
        if (effects.HasProjection) {
            maximumPixels = Math.Min(maximumPixels, _context.RemainingImageProcessingPixels / 2);
            if (maximumPixels < 1) throw new InvalidDataException("Publisher image processing pixel limit exceeded.");
        }
        OfficeRasterImage raster = metafile ? DecodeMetafile(image, objectId, maximumPixels)
            : DecodeEffectRaster(image, objectId, maximumPixels);
        if (effects.HasProjection) {
            _context.AccountImageProcessingPixels(checked((long)raster.Width * raster.Height * 2));
            raster = effects.Apply(raster, _context.Token);
            _context.Add("PUB_PICTURE_EFFECT_APPROXIMATED",
                "Native picture controls use an exact source-RGB transparency key, encoded-RGB tone adjustment and BT.709 grayscale or a 50-percent two-color threshold. Native color space, effect ordering and Publisher-rendered pixels are not qualified.",
                OfficeConversionLossKind.Approximation, PublisherEscherReader.ShapeLocation(objectId));
        }
        contentType = "image/png";
        return OfficeRasterImageEncoder.Encode(raster, OfficeImageExportFormat.Png, null, _context.Options.MaximumImageBytes, _context.Token);
    }

    private OfficeRasterImage DecodeEffectRaster(PublisherImage image, uint objectId, long maximumPixels) {
        var options = new OfficeRasterDecodeOptions {
            MaximumDecodedPixels = Math.Min(maximumPixels, 50_000_000),
            MaximumEncodedBytes = Math.Min(_context.Options.MaximumImageBytes, 128 * 1024 * 1024),
            ImageCodec = _context.Options.ImageCodec, CancellationToken = _context.Token
        };
        if (!OfficeRasterImageDecoder.TryDecode(image.Bytes, options, out OfficeRasterImage? raster, out OfficeRasterDecodeInfo info)
            || raster == null)
            throw new InvalidDataException("Publisher picture effects require a valid raster within the configured pixel and codec bounds. " + info.Diagnostic);
        string location = PublisherEscherReader.ShapeLocation(objectId);
        if (info.FramesOrPagesDiscarded)
            _context.Add(OfficeImageExportDiagnosticCodes.SourceImageStaticFrameSelected,
                "Picture-effect projection retains the selected first image frame; remaining frames or pages are omitted from the scene. Images retains the complete native payload.",
                OfficeConversionLossKind.Omission, location);
        if (info.OrientationNormalized)
            _context.Add("PUB_PICTURE_ORIENTATION_NORMALIZED", "Picture-effect decoding applies the encoded image orientation before native cropping. Publisher's orientation and crop interaction is not qualified.",
                OfficeConversionLossKind.Approximation, location);
        if (info.UsedCallerCodec)
            _context.Add(OfficeImageExportDiagnosticCodes.SourceImageDecodedByCallerCodec, "The application image codec decoded the raster for picture-effect projection.",
                OfficeConversionLossKind.None, location);
        return raster;
    }

    private OfficeRasterImage DecodeMetafile(PublisherImage image, uint objectId, long maximumPixels) {
        var diagnostics = new List<OfficeImageExportDiagnostic>();
        var codec = new OfficeRasterImageFallbackCodec(_context.Options.ImageCodec, diagnostics, PublisherEscherReader.ShapeLocation(objectId));
        _context.Token.ThrowIfCancellationRequested();
        codec.TryDecode(image.Bytes, image.ContentType, out OfficeRasterImage? raster);
        _context.Token.ThrowIfCancellationRequested();
        if (raster == null || (long)raster.Width * raster.Height > maximumPixels)
            throw new InvalidDataException("Publisher application image codec exceeded the pixel limit or returned no image.");
        foreach (OfficeImageExportDiagnostic diagnostic in diagnostics) _context.Add(diagnostic.Code, diagnostic.Message, diagnostic.LossKind, diagnostic.Source);
        if (!diagnostics.Any(item => item.LossKind == OfficeConversionLossKind.Omission))
            _context.Add("PUB_METAFILE_RASTERIZED", "A WMF/EMF picture was rasterized by the application codec. Images retains the original native payload.",
                OfficeConversionLossKind.Approximation, PublisherEscherReader.ShapeLocation(objectId));
        return raster;
    }
}
