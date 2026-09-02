using System.Globalization;
using System.Text;
using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

/// <summary>Builds a stable fingerprint of the rendering semantics for an image placement.</summary>
internal static class PdfRedactionImageIdentity {
    private const int MaximumDepth = 64;
    private const int MaximumNodes = 16384;

    internal static void Append(
        StringBuilder identity,
        PdfImagePlacement placement,
        PdfStream? imageStream,
        Dictionary<int, PdfIndirectObject> objects) {
        identity.Append(":render:")
            .Append(placement.Opacity.ToString("R", CultureInfo.InvariantCulture))
            .Append(':').Append((int)placement.EffectiveBlendMode)
            .Append(':').Append(placement.HasUnsupportedBlendMode ? '1' : '0')
            .Append(':').Append(placement.HasSoftMask ? '1' : '0')
            .Append(':').Append((int)placement.RenderingIntent)
            .Append(':').Append(placement.HasAuthoredRenderingIntent ? '1' : '0')
            .Append(':').Append(placement.RequireExactProjection ? '1' : '0');
        AppendColor(identity, placement.ImageMaskColor);
        AppendClip(identity, placement.ClipPath);
        AppendEffectiveColorSpace(identity, placement, imageStream, objects);
        AppendFillPattern(identity, placement, objects);

        identity.Append(":stream:");
        if (imageStream is null) {
            identity.Append("none");
            return;
        }

        AppendHashedObjectGraph(identity, imageStream, objects);
    }

    internal static void AppendClip(StringBuilder identity, PdfPageClipPath? clipPath) {
        if (!clipPath.HasValue) {
            identity.Append(":clip:none");
            return;
        }

        PdfPageClipPath clip = clipPath.Value;
        identity.Append(":clip:")
            .Append(clip.X.ToString("R", CultureInfo.InvariantCulture)).Append(',')
            .Append(clip.Y.ToString("R", CultureInfo.InvariantCulture)).Append(',')
            .Append(clip.Width.ToString("R", CultureInfo.InvariantCulture)).Append(',')
            .Append(clip.Height.ToString("R", CultureInfo.InvariantCulture)).Append(',')
            .Append(clip.IsRectangle ? '1' : '0').Append(',')
            .Append((int)clip.FillRule).Append(',')
            .Append(clip.IsExact ? '1' : '0').Append(',')
            .Append(clip.ContainsTextClipping ? '1' : '0').Append(',')
            .Append(clip.Commands.Count);
        for (int i = 0; i < clip.Commands.Count; i++) {
            OfficePathCommand command = clip.Commands[i];
            identity.Append(';').Append((int)command.Kind).Append(',')
                .Append(command.Point.X.ToString("R", CultureInfo.InvariantCulture)).Append(',')
                .Append(command.Point.Y.ToString("R", CultureInfo.InvariantCulture)).Append(',')
                .Append(command.ControlPoint1.X.ToString("R", CultureInfo.InvariantCulture)).Append(',')
                .Append(command.ControlPoint1.Y.ToString("R", CultureInfo.InvariantCulture)).Append(',')
                .Append(command.ControlPoint2.X.ToString("R", CultureInfo.InvariantCulture)).Append(',')
                .Append(command.ControlPoint2.Y.ToString("R", CultureInfo.InvariantCulture));
        }
    }

    internal static void AppendObjectGraph(
        StringBuilder identity,
        PdfObject? value,
        Dictionary<int, PdfIndirectObject> objects) {
        if (value is null) {
            identity.Append("none");
            return;
        }

        AppendHashedObjectGraph(identity, value, objects);
    }

    private static void AppendColor(StringBuilder identity, OfficeColor color) =>
        identity.Append(":mask-color:").Append(color.R).Append(',').Append(color.G).Append(',').Append(color.B).Append(',').Append(color.A);

    private static void AppendHashedObjectGraph(
        StringBuilder identity,
        PdfObject value,
        Dictionary<int, PdfIndirectObject> objects) {
        using var fingerprint = new PdfObjectGraphFingerprint(
            objects,
            MaximumDepth,
            MaximumNodes,
            includeDictionaryKey: static key => !string.Equals(key, "Length", StringComparison.Ordinal));
        fingerprint.AppendRoot(value);
        AppendText(identity, 'H', Convert.ToBase64String(fingerprint.Complete()));
    }

    private static void AppendEffectiveColorSpace(
        StringBuilder identity,
        PdfImagePlacement placement,
        PdfStream? imageStream,
        Dictionary<int, PdfIndirectObject> objects) {
        if (imageStream is null) return;
        PdfObject? declaration = imageStream.Dictionary.Items.TryGetValue("ColorSpace", out PdfObject? colorSpace)
            ? colorSpace
            : imageStream.Dictionary.Items.TryGetValue("CS", out PdfObject? abbreviatedColorSpace)
                ? abbreviatedColorSpace
                : null;
        if (PdfObjectLookup.ResolveChain(objects, declaration) is not PdfName name || IsBuiltInColorSpace(name.Name)) return;

        PdfDictionary? resources = placement.InlineImageResources ?? placement.EffectiveResources;
        if (!TryGetResource(resources, "ColorSpace", name.Name, objects, out PdfObject? resource)) return;
        identity.Append(":effective-color-space:");
        AppendHashedObjectGraph(identity, resource, objects);
    }

    private static void AppendFillPattern(
        StringBuilder identity,
        PdfImagePlacement placement,
        Dictionary<int, PdfIndirectObject> objects) {
        if (!placement.FillPattern.HasValue) return;
        PdfPagePatternSelection pattern = placement.FillPattern.Value;
        identity.Append(":fill-pattern:");
        AppendText(identity, 'N', pattern.Name);
        identity.Append(':').Append(pattern.ComponentCount)
            .Append(':').Append((int)pattern.RenderingIntent)
            .Append(':').Append(pattern.BaseColorSpace.HasValue ? (int)pattern.BaseColorSpace.Value.Kind : -1);
        if (pattern.Tint.HasValue) AppendColor(identity, pattern.Tint.Value);
        Matrix2D transform = pattern.PaintTransform;
        identity.Append(':').Append(transform.A.ToString("R", CultureInfo.InvariantCulture)).Append(',')
            .Append(transform.B.ToString("R", CultureInfo.InvariantCulture)).Append(',')
            .Append(transform.C.ToString("R", CultureInfo.InvariantCulture)).Append(',')
            .Append(transform.D.ToString("R", CultureInfo.InvariantCulture)).Append(',')
            .Append(transform.E.ToString("R", CultureInfo.InvariantCulture)).Append(',')
            .Append(transform.F.ToString("R", CultureInfo.InvariantCulture));

        if (!TryGetResource(placement.EffectiveResources, "Pattern", pattern.Name, objects, out PdfObject? resource)) return;
        AppendHashedObjectGraph(identity, resource, objects);
    }

    private static bool TryGetResource(
        PdfDictionary? resources,
        string category,
        string name,
        Dictionary<int, PdfIndirectObject> objects,
        out PdfObject resource) {
        resource = null!;
        if (resources is null ||
            !resources.Items.TryGetValue(category, out PdfObject? categoryObject) ||
            PdfObjectLookup.ResolveChain(objects, categoryObject) is not PdfDictionary categoryDictionary ||
            !categoryDictionary.Items.TryGetValue(name, out PdfObject? value)) return false;
        resource = value;
        return true;
    }

    private static bool IsBuiltInColorSpace(string name) =>
        name is "DeviceGray" or "DeviceRGB" or "DeviceCMYK" or "Pattern" or "G" or "RGB" or "CMYK" or "I";

    private static void AppendText(StringBuilder identity, char prefix, string value) =>
        identity.Append(prefix).Append(value.Length).Append(':').Append(value);

}
