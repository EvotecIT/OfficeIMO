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
        bool hasCrop = sourceLeft > 0D || sourceTop > 0D || sourceWidth < 1D || sourceHeight < 1D;
        if (hasCrop && string.IsNullOrEmpty(clipPathId)) {
            throw new ArgumentException("A clip path identifier is required when source crop is used.", nameof(clipPathId));
        }

        double imageX = x;
        double imageY = y;
        double imageWidth = width;
        double imageHeight = height;
        if (hasCrop) {
            imageWidth = width / sourceWidth;
            imageHeight = height / sourceHeight;
            imageX = x - (sourceLeft * imageWidth);
            imageY = y - (sourceTop * imageHeight);
        }

        string? transform = BuildTransform(x, y, width, height, rotationDegrees, flipHorizontal, flipVertical);
        bool transformCroppedImage = hasCrop && transform != null;
        if (!string.IsNullOrEmpty(clipPathId)) {
            builder.AppendRectClipPathDefinition(clipPathId!, clipX, clipY, clipWidth, clipHeight);
        }

        if (transformCroppedImage) {
            builder.Append("<g")
                .AppendClipPathReference(clipPathId!)
                .AppendAttribute("transform", transform)
                .Append(">");
        }

        builder.Append("<image")
            .AppendNumberAttribute("x", imageX)
            .AppendNumberAttribute("y", imageY)
            .AppendNumberAttribute("width", imageWidth)
            .AppendNumberAttribute("height", imageHeight);
        if (!transformCroppedImage) {
            if (!string.IsNullOrEmpty(clipPathId)) {
                builder.AppendClipPathReference(clipPathId!);
            }

            if (transform != null) {
                builder.AppendAttribute("transform", transform);
            }
        }

        if (!string.IsNullOrWhiteSpace(preserveAspectRatio)) {
            builder.AppendAttribute("preserveAspectRatio", preserveAspectRatio!);
        }

        builder.AppendAttribute("href", href).Append("/>");
        if (transformCroppedImage) {
            builder.Append("</g>");
        }

        return builder;
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
        if (writer == null) {
            throw new ArgumentNullException(nameof(writer));
        }

        if (string.IsNullOrEmpty(href) || width <= 0D || height <= 0D) {
            return;
        }

        writer.WriteStartElement("image", svgNamespace);
        writeAdditionalAttributes?.Invoke(writer);
        writer.WriteNumberAttribute("x", x);
        writer.WriteNumberAttribute("y", y);
        writer.WriteNumberAttribute("width", width);
        writer.WriteNumberAttribute("height", height);
        if (!string.IsNullOrWhiteSpace(preserveAspectRatio)) {
            writer.WriteAttributeString("preserveAspectRatio", preserveAspectRatio);
        }

        string? transform = BuildTransform(
            x,
            y,
            width,
            height,
            rotationDegrees,
            flipHorizontal,
            flipVertical,
            rotationCenterX,
            rotationCenterY);
        if (transform != null) {
            writer.WriteAttributeString("transform", transform);
        }

        writer.WriteAttributeString("href", href);
        writer.WriteEndElement();
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

    private static string? BuildTransform(
        double x,
        double y,
        double width,
        double height,
        double rotationDegrees,
        bool flipHorizontal,
        bool flipVertical,
        double? rotationCenterX = null,
        double? rotationCenterY = null) {
        bool hasRotation = Math.Abs(rotationDegrees) >= 0.000001D;
        bool hasFlip = flipHorizontal || flipVertical;
        if (!hasRotation && !hasFlip) {
            return null;
        }

        double centerX = rotationCenterX ?? x + (width / 2D);
        double centerY = rotationCenterY ?? y + (height / 2D);
        if (!hasFlip) {
            return OfficeSvgFormatting.FormatRotateTransform(rotationDegrees, centerX, centerY);
        }

        double scaleX = flipHorizontal ? -1D : 1D;
        double scaleY = flipVertical ? -1D : 1D;
        var transform = new StringBuilder();
        transform.Append("translate(")
            .Append(OfficeSvgFormatting.FormatNumber(centerX))
            .Append(' ')
            .Append(OfficeSvgFormatting.FormatNumber(centerY))
            .Append(')');
        if (hasRotation) {
            transform.Append(' ').Append(OfficeSvgFormatting.FormatRotateTransform(rotationDegrees));
        }

        transform.Append(" scale(")
            .Append(OfficeSvgFormatting.FormatNumber(scaleX))
            .Append(' ')
            .Append(OfficeSvgFormatting.FormatNumber(scaleY))
            .Append(')');
        transform.Append(" translate(")
            .Append(OfficeSvgFormatting.FormatNumber(-centerX))
            .Append(' ')
            .Append(OfficeSvgFormatting.FormatNumber(-centerY))
            .Append(')');
        return transform.ToString();
    }
}
