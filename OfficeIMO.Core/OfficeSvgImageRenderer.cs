using System;
using System.Text;
using System.Xml;

namespace OfficeIMO.Drawing;

/// <summary>
/// Shared dependency-free SVG renderer for projected bitmap/vector images.
/// </summary>
public static class OfficeSvgImageRenderer {
    /// <summary>
    /// Appends an SVG image using normalized source cropping, optional clipping, rotation, and flips.
    /// </summary>
    /// <param name="builder">Markup builder.</param>
    /// <param name="href">Resolved SVG image reference, such as a data URI.</param>
    /// <param name="x">Destination left coordinate.</param>
    /// <param name="y">Destination top coordinate.</param>
    /// <param name="width">Destination width.</param>
    /// <param name="height">Destination height.</param>
    /// <param name="clipPathId">Optional clip-path identifier. Required when source crop should be clipped to the destination rectangle.</param>
    /// <param name="clipX">Clip rectangle left coordinate.</param>
    /// <param name="clipY">Clip rectangle top coordinate.</param>
    /// <param name="clipWidth">Clip rectangle width.</param>
    /// <param name="clipHeight">Clip rectangle height.</param>
    /// <param name="sourceLeft">Normalized source left crop ratio.</param>
    /// <param name="sourceTop">Normalized source top crop ratio.</param>
    /// <param name="sourceWidth">Normalized visible source width ratio.</param>
    /// <param name="sourceHeight">Normalized visible source height ratio.</param>
    /// <param name="rotationDegrees">Clockwise SVG rotation in degrees.</param>
    /// <param name="flipHorizontal">Whether to mirror the image horizontally around the destination center.</param>
    /// <param name="flipVertical">Whether to mirror the image vertically around the destination center.</param>
    /// <param name="preserveAspectRatio">Optional SVG preserveAspectRatio value.</param>
    /// <returns>The supplied builder for call chaining.</returns>
    public static StringBuilder AppendImage(
        StringBuilder builder,
        string href,
        double x,
        double y,
        double width,
        double height,
        string? clipPathId = null,
        double clipX = 0D,
        double clipY = 0D,
        double clipWidth = 0D,
        double clipHeight = 0D,
        double sourceLeft = 0D,
        double sourceTop = 0D,
        double sourceWidth = 1D,
        double sourceHeight = 1D,
        double rotationDegrees = 0D,
        bool flipHorizontal = false,
        bool flipVertical = false,
        string? preserveAspectRatio = null) {
        if (builder == null) {
            throw new ArgumentNullException(nameof(builder));
        }

        if (string.IsNullOrEmpty(href) || width <= 0D || height <= 0D) {
            return builder;
        }

        sourceWidth = Math.Max(0.001D, sourceWidth);
        sourceHeight = Math.Max(0.001D, sourceHeight);
        OfficeImageSourceCrop sourceCrop = OfficeImageSourceCrop.FromClampedFractions(
            sourceLeft,
            sourceTop,
            Math.Max(0D, 1D - sourceLeft - sourceWidth),
            Math.Max(0D, 1D - sourceTop - sourceHeight));
        OfficeImagePlacement? clipPlacement = null;
        if (!string.IsNullOrEmpty(clipPathId) && clipWidth > 0D && clipHeight > 0D) {
            clipPlacement = new OfficeImagePlacement(clipX, clipY, clipWidth, clipHeight);
        }

        return AppendImage(
            builder,
            href,
            new OfficeImageProjection(
                new OfficeImagePlacement(x, y, width, height),
                sourceCrop,
                rotationDegrees,
                flipHorizontal: flipHorizontal,
                flipVertical: flipVertical),
            clipPathId,
            clipPlacement,
            preserveAspectRatio);
    }

    /// <summary>
    /// Appends an SVG image using a shared projection that carries placement, source crop, rotation, and flips.
    /// </summary>
    /// <param name="builder">Markup builder.</param>
    /// <param name="href">Resolved SVG image reference, such as a data URI.</param>
    /// <param name="projection">Shared image projection.</param>
    /// <param name="clipPathId">Optional clip-path identifier. Required when the projection has a source crop.</param>
    /// <param name="clipRectangle">Optional clip rectangle. Defaults to the projection placement when source crop is used.</param>
    /// <param name="preserveAspectRatio">Optional SVG preserveAspectRatio value.</param>
    /// <returns>The supplied builder for call chaining.</returns>
    public static StringBuilder AppendImage(
        StringBuilder builder,
        string href,
        OfficeImageProjection projection,
        string? clipPathId = null,
        OfficeImagePlacement? clipRectangle = null,
        string? preserveAspectRatio = null) {
        if (builder == null) {
            throw new ArgumentNullException(nameof(builder));
        }

        if (string.IsNullOrEmpty(href) || projection.Width <= 0D || projection.Height <= 0D) {
            return builder;
        }

        if (projection.HasCrop && string.IsNullOrEmpty(clipPathId)) {
            throw new ArgumentException("A clip path identifier is required when source crop is used.", nameof(clipPathId));
        }

        SvgImageLayout layout = CreateLayout(projection, clipRectangle);

        if (!string.IsNullOrEmpty(clipPathId) && layout.EffectiveClip != null) {
            OfficeImagePlacement clip = layout.EffectiveClip.Value;
            builder.AppendRectClipPathDefinition(clipPathId!, clip.X, clip.Y, clip.Width, clip.Height);
        }

        if (layout.TransformCroppedImage) {
            builder.Append("<g")
                .AppendClipPathReference(clipPathId!)
                .AppendAttribute("transform", layout.Transform)
                .Append(">");
        }

        builder.Append("<image")
            .AppendNumberAttribute("x", layout.ImagePlacement.X)
            .AppendNumberAttribute("y", layout.ImagePlacement.Y)
            .AppendNumberAttribute("width", layout.ImagePlacement.Width)
            .AppendNumberAttribute("height", layout.ImagePlacement.Height);
        if (!layout.TransformCroppedImage) {
            if (!string.IsNullOrEmpty(clipPathId) && layout.EffectiveClip != null) {
                builder.AppendClipPathReference(clipPathId!);
            }

            if (layout.Transform != null) {
                builder.AppendAttribute("transform", layout.Transform);
            }
        }

        if (!string.IsNullOrWhiteSpace(preserveAspectRatio)) {
            builder.AppendAttribute("preserveAspectRatio", preserveAspectRatio!);
        }

        builder.AppendAttribute("href", href).Append("/>");
        if (layout.TransformCroppedImage) {
            builder.Append("</g>");
        }

        return builder;
    }

    /// <summary>
    /// Appends an SVG image inside a viewport, using the projection placement as the clip for source-cropped images
    /// and the viewport as the clip for uncropped images.
    /// </summary>
    /// <param name="builder">Markup builder.</param>
    /// <param name="href">Resolved SVG image reference, such as a data URI.</param>
    /// <param name="projection">Shared image projection.</param>
    /// <param name="clipPathId">Clip-path identifier used for the viewport or crop clip.</param>
    /// <param name="viewport">Viewport rectangle that clips uncropped images.</param>
    /// <param name="preserveAspectRatio">Optional SVG preserveAspectRatio value.</param>
    /// <returns>The supplied builder for call chaining.</returns>
    public static StringBuilder AppendImageInViewport(
        StringBuilder builder,
        string href,
        OfficeImageProjection projection,
        string clipPathId,
        OfficeImagePlacement viewport,
        string? preserveAspectRatio = null) {
        if (string.IsNullOrEmpty(clipPathId)) {
            throw new ArgumentException("A clip path identifier is required for viewport image rendering.", nameof(clipPathId));
        }

        return AppendImage(
            builder,
            href,
            projection,
            clipPathId,
            projection.HasCrop ? projection.Placement : viewport,
            preserveAspectRatio);
    }

    /// <summary>
    /// Writes an SVG image element using shared numeric formatting, optional preserve-aspect behavior, rotation, and flips.
    /// </summary>
    /// <param name="writer">SVG writer.</param>
    /// <param name="svgNamespace">SVG namespace URI.</param>
    /// <param name="href">Resolved SVG image reference, such as a data URI.</param>
    /// <param name="x">Destination left coordinate.</param>
    /// <param name="y">Destination top coordinate.</param>
    /// <param name="width">Destination width.</param>
    /// <param name="height">Destination height.</param>
    /// <param name="rotationDegrees">Clockwise SVG rotation in degrees.</param>
    /// <param name="rotationCenterX">Rotation center X coordinate. When omitted, the destination center is used.</param>
    /// <param name="rotationCenterY">Rotation center Y coordinate. When omitted, the destination center is used.</param>
    /// <param name="flipHorizontal">Whether to mirror the image horizontally around the rotation center.</param>
    /// <param name="flipVertical">Whether to mirror the image vertically around the rotation center.</param>
    /// <param name="preserveAspectRatio">Optional SVG preserveAspectRatio value.</param>
    /// <param name="writeAdditionalAttributes">Optional callback for adapter-specific attributes.</param>
    public static void WriteImage(
        XmlWriter writer,
        string svgNamespace,
        string href,
        double x,
        double y,
        double width,
        double height,
        double rotationDegrees = 0D,
        double? rotationCenterX = null,
        double? rotationCenterY = null,
        bool flipHorizontal = false,
        bool flipVertical = false,
        string? preserveAspectRatio = null,
        Action<XmlWriter>? writeAdditionalAttributes = null) {
        WriteImage(
            writer,
            svgNamespace,
            href,
            new OfficeImageProjection(
                new OfficeImagePlacement(x, y, width, height),
                rotationDegrees: rotationDegrees,
                rotationCenterX: rotationCenterX,
                rotationCenterY: rotationCenterY,
                flipHorizontal: flipHorizontal,
                flipVertical: flipVertical),
            preserveAspectRatio,
            writeAdditionalAttributes);
    }

    /// <summary>
    /// Writes an SVG image element using a shared projection that carries placement, rotation, and flips.
    /// </summary>
    /// <param name="writer">SVG writer.</param>
    /// <param name="svgNamespace">SVG namespace URI.</param>
    /// <param name="href">Resolved SVG image reference, such as a data URI.</param>
    /// <param name="projection">Shared image projection.</param>
    /// <param name="preserveAspectRatio">Optional SVG preserveAspectRatio value.</param>
    /// <param name="writeAdditionalAttributes">Optional callback for adapter-specific attributes.</param>
    /// <param name="clipPathId">Optional clip-path identifier. Required when the projection has a source crop.</param>
    /// <param name="clipRectangle">Optional clip rectangle. Defaults to the projection placement when source crop is used.</param>
    public static void WriteImage(
        XmlWriter writer,
        string svgNamespace,
        string href,
        OfficeImageProjection projection,
        string? preserveAspectRatio = null,
        Action<XmlWriter>? writeAdditionalAttributes = null,
        string? clipPathId = null,
        OfficeImagePlacement? clipRectangle = null) {
        if (writer == null) {
            throw new ArgumentNullException(nameof(writer));
        }

        if (string.IsNullOrEmpty(href) || projection.Width <= 0D || projection.Height <= 0D) {
            return;
        }

        if (projection.HasCrop && string.IsNullOrEmpty(clipPathId)) {
            throw new ArgumentException("A clip path identifier is required when source crop is used.", nameof(clipPathId));
        }

        SvgImageLayout layout = CreateLayout(projection, clipRectangle);
        if (!string.IsNullOrEmpty(clipPathId) && layout.EffectiveClip != null) {
            WriteRectClipPathDefinition(writer, svgNamespace, clipPathId!, layout.EffectiveClip.Value);
        }

        if (layout.TransformCroppedImage) {
            writer.WriteStartElement("g", svgNamespace);
            writer.WriteAttributeString("clip-path", "url(#" + clipPathId + ")");
            writer.WriteAttributeString("transform", layout.Transform);
        }

        writer.WriteStartElement("image", svgNamespace);
        writeAdditionalAttributes?.Invoke(writer);
        writer.WriteNumberAttribute("x", layout.ImagePlacement.X);
        writer.WriteNumberAttribute("y", layout.ImagePlacement.Y);
        writer.WriteNumberAttribute("width", layout.ImagePlacement.Width);
        writer.WriteNumberAttribute("height", layout.ImagePlacement.Height);
        if (!layout.TransformCroppedImage) {
            if (!string.IsNullOrEmpty(clipPathId) && layout.EffectiveClip != null) {
                writer.WriteAttributeString("clip-path", "url(#" + clipPathId + ")");
            }

            if (layout.Transform != null) {
                writer.WriteAttributeString("transform", layout.Transform);
            }
        }

        if (!string.IsNullOrWhiteSpace(preserveAspectRatio)) {
            writer.WriteAttributeString("preserveAspectRatio", preserveAspectRatio);
        }

        writer.WriteAttributeString("href", href);
        writer.WriteEndElement();
        if (layout.TransformCroppedImage) {
            writer.WriteEndElement();
        }
    }

    /// <summary>
    /// Writes an SVG image inside a viewport, using the projection placement as the clip for source-cropped images
    /// and the viewport as the clip for uncropped images.
    /// </summary>
    /// <param name="writer">SVG writer.</param>
    /// <param name="svgNamespace">SVG namespace URI.</param>
    /// <param name="href">Resolved SVG image reference, such as a data URI.</param>
    /// <param name="projection">Shared image projection.</param>
    /// <param name="clipPathId">Clip-path identifier used for the viewport or crop clip.</param>
    /// <param name="viewport">Viewport rectangle that clips uncropped images.</param>
    /// <param name="preserveAspectRatio">Optional SVG preserveAspectRatio value.</param>
    /// <param name="writeAdditionalAttributes">Optional callback for adapter-specific attributes.</param>
    public static void WriteImageInViewport(
        XmlWriter writer,
        string svgNamespace,
        string href,
        OfficeImageProjection projection,
        string clipPathId,
        OfficeImagePlacement viewport,
        string? preserveAspectRatio = null,
        Action<XmlWriter>? writeAdditionalAttributes = null) {
        if (string.IsNullOrEmpty(clipPathId)) {
            throw new ArgumentException("A clip path identifier is required for viewport image rendering.", nameof(clipPathId));
        }

        WriteImage(
            writer,
            svgNamespace,
            href,
            projection,
            preserveAspectRatio,
            writeAdditionalAttributes,
            clipPathId,
            projection.HasCrop ? projection.Placement : viewport);
    }

    /// <summary>
    /// Builds a data URI image reference from a known-safe content type and image bytes.
    /// </summary>
    /// <param name="contentType">Image MIME content type.</param>
    /// <param name="bytes">Image bytes.</param>
    /// <returns>SVG data URI image reference.</returns>
    public static string CreateDataUri(string contentType, byte[] bytes) {
        if (string.IsNullOrWhiteSpace(contentType)) {
            throw new ArgumentException("Content type is required.", nameof(contentType));
        }

        if (bytes == null) {
            throw new ArgumentNullException(nameof(bytes));
        }

        return "data:" + contentType + ";base64," + Convert.ToBase64String(bytes);
    }

    /// <summary>
    /// Creates an SVG-safe data URI for image bytes that can either be embedded directly or transcoded through the shared raster decoder.
    /// </summary>
    /// <param name="declaredContentType">Optional content type from the source package or caller.</param>
    /// <param name="bytes">Image bytes.</param>
    /// <param name="fileName">Optional file name or extension used when metadata is absent or generic.</param>
    /// <param name="dataUri">SVG data URI when the image can be represented in SVG output.</param>
    /// <returns><see langword="true" /> when the image can be embedded in SVG output.</returns>
    public static bool TryCreateDataUri(string? declaredContentType, byte[]? bytes, string? fileName, out string dataUri) {
        return TryCreateDataUri(declaredContentType, bytes, fileName, null, out dataUri);
    }

    /// <summary>
    /// Creates an SVG-safe data URI, using an optional shared raster codec when Drawing cannot decode the source format.
    /// </summary>
    public static bool TryCreateDataUri(
        string? declaredContentType,
        byte[]? bytes,
        string? fileName,
        IOfficeRasterImageCodec? imageCodec,
        out string dataUri) {
        dataUri = string.Empty;
        if (bytes == null || bytes.Length == 0) {
            return false;
        }

        if (TryResolveEmbeddableContentType(declaredContentType, bytes, fileName, out string contentType)) {
            dataUri = CreateDataUri(contentType, bytes);
            return true;
        }

        if (OfficeRasterImageDecoder.TryDecode(bytes, out OfficeRasterImage? raster) && raster != null) {
            dataUri = CreateDataUri("image/png", OfficePngWriter.Encode(raster));
            return true;
        }

        if (imageCodec != null && imageCodec.TryDecode((byte[])bytes.Clone(), declaredContentType, out raster) && raster != null) {
            dataUri = CreateDataUri("image/png", OfficePngWriter.Encode(raster));
            return true;
        }

        return false;
    }

    /// <summary>
    /// Resolves the MIME content type for image formats that can be embedded directly in SVG image elements.
    /// </summary>
    /// <param name="format">Detected image format.</param>
    /// <param name="contentType">Resolved MIME content type when supported.</param>
    /// <returns><see langword="true" /> when the format can be embedded as an SVG image href.</returns>
    public static bool TryGetEmbeddableContentType(OfficeImageFormat format, out string contentType) {
        switch (format) {
            case OfficeImageFormat.Png:
            case OfficeImageFormat.Jpeg:
            case OfficeImageFormat.Gif:
            case OfficeImageFormat.Svg:
            case OfficeImageFormat.Webp:
                contentType = OfficeImageInfo.GetMimeType(format);
                return true;
            default:
                contentType = string.Empty;
                return false;
        }
    }

    /// <summary>
    /// Resolves and normalizes a MIME content type when it can be embedded directly in an SVG image element.
    /// </summary>
    /// <param name="contentType">MIME content type, optionally with parameters.</param>
    /// <param name="normalizedContentType">Canonical MIME content type when supported.</param>
    /// <returns><see langword="true" /> when the content type can be embedded as an SVG image href.</returns>
    public static bool TryGetEmbeddableContentType(string? contentType, out string normalizedContentType) {
        OfficeImageFormat format = OfficeImageInfo.FromMimeType(contentType);
        return TryGetEmbeddableContentType(format, out normalizedContentType);
    }

    /// <summary>
    /// Resolves an SVG-embeddable image content type from declared package metadata, image bytes, or a file name.
    /// </summary>
    /// <param name="declaredContentType">Optional content type from the source package or caller.</param>
    /// <param name="bytes">Optional image bytes used for dependency-free signature sniffing.</param>
    /// <param name="fileName">Optional file name or extension used when metadata is absent or generic.</param>
    /// <param name="contentType">Resolved MIME content type when supported.</param>
    /// <returns><see langword="true" /> when the image can be embedded directly in an SVG image href.</returns>
    public static bool TryResolveEmbeddableContentType(string? declaredContentType, byte[]? bytes, string? fileName, out string contentType) {
        string normalized = NormalizeContentType(declaredContentType);
        if (!string.IsNullOrEmpty(normalized) && !IsGenericContentType(normalized)) {
            if (TryGetEmbeddableContentType(normalized, out string declaredFormatContentType)) {
                contentType = declaredFormatContentType;
                return true;
            }

            contentType = string.Empty;
            return false;
        }

        if (TrySniffEmbeddableFormat(bytes, out OfficeImageFormat sniffedFormat) &&
            TryGetEmbeddableContentType(sniffedFormat, out contentType)) {
            return true;
        }

        OfficeImageFormat extensionFormat = OfficeImageReader.FromExtension(fileName);
        if (TryGetEmbeddableContentType(extensionFormat, out contentType)) {
            return true;
        }

        contentType = string.Empty;
        return false;
    }

    private static SvgImageLayout CreateLayout(OfficeImageProjection projection, OfficeImagePlacement? clipRectangle) {
        double imageX = projection.X;
        double imageY = projection.Y;
        double imageWidth = projection.Width;
        double imageHeight = projection.Height;
        if (projection.HasCrop) {
            imageWidth = projection.Width / projection.SourceWidth;
            imageHeight = projection.Height / projection.SourceHeight;
            imageX = projection.X - (projection.SourceLeft * imageWidth);
            imageY = projection.Y - (projection.SourceTop * imageHeight);
        }

        string? transform = OfficeSvgFormatting.FormatImageFrameTransform(projection.CreateFrameTransform());
        OfficeImagePlacement? effectiveClip = clipRectangle;
        if (projection.HasCrop && effectiveClip == null) {
            effectiveClip = projection.Placement;
        }

        return new SvgImageLayout(
            new OfficeImagePlacement(imageX, imageY, imageWidth, imageHeight),
            effectiveClip,
            transform,
            projection.HasCrop && transform != null);
    }

    private static void WriteRectClipPathDefinition(XmlWriter writer, string svgNamespace, string clipPathId, OfficeImagePlacement clip) {
        writer.WriteStartElement("clipPath", svgNamespace);
        writer.WriteAttributeString("id", clipPathId);
        writer.WriteStartElement("rect", svgNamespace);
        writer.WriteNumberAttribute("x", clip.X);
        writer.WriteNumberAttribute("y", clip.Y);
        writer.WriteNumberAttribute("width", clip.Width);
        writer.WriteNumberAttribute("height", clip.Height);
        writer.WriteEndElement();
        writer.WriteEndElement();
    }

    private readonly struct SvgImageLayout {
        internal SvgImageLayout(OfficeImagePlacement imagePlacement, OfficeImagePlacement? effectiveClip, string? transform, bool transformCroppedImage) {
            ImagePlacement = imagePlacement;
            EffectiveClip = effectiveClip;
            Transform = transform;
            TransformCroppedImage = transformCroppedImage;
        }

        internal OfficeImagePlacement ImagePlacement { get; }

        internal OfficeImagePlacement? EffectiveClip { get; }

        internal string? Transform { get; }

        internal bool TransformCroppedImage { get; }
    }

    private static string NormalizeContentType(string? contentType) {
        return OfficeImageInfo.NormalizeMimeType(contentType);
    }

    private static bool IsGenericContentType(string contentType) =>
        string.Equals(contentType, "application/octet-stream", StringComparison.OrdinalIgnoreCase) ||
        string.Equals(contentType, "binary/octet-stream", StringComparison.OrdinalIgnoreCase);

    private static bool TrySniffEmbeddableFormat(byte[]? data, out OfficeImageFormat format) {
        format = OfficeImageFormat.Unknown;
        if (data == null || data.Length == 0) {
            return false;
        }

        if (data.Length >= 8 &&
            data[0] == 0x89 &&
            data[1] == (byte)'P' &&
            data[2] == (byte)'N' &&
            data[3] == (byte)'G' &&
            data[4] == 0x0D &&
            data[5] == 0x0A &&
            data[6] == 0x1A &&
            data[7] == 0x0A) {
            format = OfficeImageFormat.Png;
            return true;
        }

        if (data.Length >= 3 &&
            data[0] == 0xFF &&
            data[1] == 0xD8 &&
            data[2] == 0xFF) {
            format = OfficeImageFormat.Jpeg;
            return true;
        }

        if (data.Length >= 6 &&
            data[0] == (byte)'G' &&
            data[1] == (byte)'I' &&
            data[2] == (byte)'F' &&
            data[3] == (byte)'8' &&
            (data[4] == (byte)'7' || data[4] == (byte)'9') &&
            data[5] == (byte)'a') {
            format = OfficeImageFormat.Gif;
            return true;
        }

        if (LooksLikeSvg(data)) {
            format = OfficeImageFormat.Svg;
            return true;
        }

        return false;
    }

    private static bool LooksLikeSvg(byte[] data) {
        int index = SkipBomAndWhitespace(data, 0);
        while (index < data.Length && data[index] == (byte)'<') {
            int tagStart = SkipAsciiWhitespace(data, index + 1);
            if (StartsWithAscii(data, tagStart, "svg")) {
                return true;
            }

            if (StartsWithAscii(data, tagStart, "!--")) {
                int commentEnd = IndexOfAscii(data, tagStart + 3, "-->");
                if (commentEnd < 0) {
                    return false;
                }

                index = SkipAsciiWhitespace(data, commentEnd + 3);
                continue;
            }

            if (StartsWithAscii(data, tagStart, "!doctype")) {
                int declarationEnd = IndexOfByte(data, tagStart + 8, (byte)'>');
                if (declarationEnd < 0) {
                    return false;
                }

                index = SkipAsciiWhitespace(data, declarationEnd + 1);
                continue;
            }

            if (tagStart < data.Length && data[tagStart] == (byte)'?') {
                int processingInstructionEnd = IndexOfAscii(data, tagStart + 1, "?>");
                if (processingInstructionEnd < 0) {
                    return false;
                }

                index = SkipAsciiWhitespace(data, processingInstructionEnd + 2);
                continue;
            }

            return false;
        }

        return false;
    }

    private static int SkipBomAndWhitespace(byte[] data, int index) {
        if (data.Length >= index + 3 &&
            data[index] == 0xEF &&
            data[index + 1] == 0xBB &&
            data[index + 2] == 0xBF) {
            index += 3;
        }

        return SkipAsciiWhitespace(data, index);
    }

    private static int SkipAsciiWhitespace(byte[] data, int index) {
        while (index < data.Length && IsAsciiWhitespace(data[index])) {
            index++;
        }

        return index;
    }

    private static int IndexOfByte(byte[] data, int startIndex, byte value) {
        for (int i = startIndex; i < data.Length; i++) {
            if (data[i] == value) {
                return i;
            }
        }

        return -1;
    }

    private static bool IsAsciiWhitespace(byte value) =>
        value == (byte)' ' ||
        value == (byte)'\t' ||
        value == (byte)'\r' ||
        value == (byte)'\n';

    private static int IndexOfAscii(byte[] data, int startIndex, string value) {
        for (int i = startIndex; i <= data.Length - value.Length; i++) {
            if (StartsWithAscii(data, i, value)) {
                return i;
            }
        }

        return -1;
    }

    private static bool StartsWithAscii(byte[] data, int startIndex, string value) {
        if (startIndex < 0 || startIndex + value.Length > data.Length) {
            return false;
        }

        for (int i = 0; i < value.Length; i++) {
            byte actual = data[startIndex + i];
            byte expected = (byte)value[i];
            if (actual >= (byte)'A' && actual <= (byte)'Z') {
                actual = (byte)(actual + 32);
            }

            if (expected >= (byte)'A' && expected <= (byte)'Z') {
                expected = (byte)(expected + 32);
            }

            if (actual != expected) {
                return false;
            }
        }

        return true;
    }
}
