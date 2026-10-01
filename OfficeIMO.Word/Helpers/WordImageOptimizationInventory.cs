using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using System.Globalization;
using System.Threading;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Drawing.Pictures;
using W = DocumentFormat.OpenXml.Wordprocessing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using V = DocumentFormat.OpenXml.Vml;

namespace OfficeIMO.Word;

/// <summary>Maps every embedded media part to its XML references and conservative source-pixel requirements.</summary>
internal sealed class WordImageOptimizationInventory {
    internal sealed class Media {
        internal Media(ImagePart part) => Part = part;
        internal ImagePart Part { get; }
        internal int References;
        internal int TargetWidth;
        internal int TargetHeight;
        internal bool UnknownPlacement;
    }

    internal readonly Dictionary<Uri, Media> Images = new();
    internal int ExternalReferences;

    internal static WordImageOptimizationInventory Build(WordprocessingDocument package,
        WordImageOptimizationOptions options, CancellationToken token) {
        var result = new WordImageOptimizationInventory();
        List<OpenXmlPart> parts = WordPackageParts.Enumerate(package, options.MaxPackageParts).ToList();
        foreach (ImagePart image in parts.OfType<ImagePart>()) result.Images.Add(image.Uri, new Media(image));
        foreach (OpenXmlPart part in parts) {
            token.ThrowIfCancellationRequested();
            if (part is ImagePart || !part.ContentType.EndsWith("xml", StringComparison.OrdinalIgnoreCase)) continue;
            OpenXmlPartRootElement? root = part.RootElement;
            if (root == null) continue;
            Dictionary<string, OpenXmlPart> relationships = part.Parts.ToDictionary(pair => pair.RelationshipId, pair => pair.OpenXmlPart);
            var external = new HashSet<string>(part.ExternalRelationships
                .Where(relationship => relationship.RelationshipType.EndsWith("/image", StringComparison.Ordinal))
                .Select(relationship => relationship.Id), StringComparer.Ordinal);
            foreach (OpenXmlElement element in root.Descendants().Prepend(root)) {
                token.ThrowIfCancellationRequested();
                foreach (OpenXmlAttribute attribute in element.GetAttributes()) {
                    if (string.IsNullOrEmpty(attribute.Value)) continue;
                    if (attribute.NamespaceUri != "http://schemas.openxmlformats.org/officeDocument/2006/relationships" &&
                        !(attribute.LocalName == "relid" && attribute.NamespaceUri == "urn:schemas-microsoft-com:office:office")) continue;
                    if (external.Contains(attribute.Value!)) { result.ExternalReferences++; continue; }
                    if (!relationships.TryGetValue(attribute.Value!, out OpenXmlPart? target) || target is not ImagePart image) continue;
                    Media media = result.Images[image.Uri];
                    media.References++;
                    if (TryGetRequiredPixels(element, options.TargetDpi, out int width, out int height)) {
                        media.TargetWidth = Math.Max(media.TargetWidth, width);
                        media.TargetHeight = Math.Max(media.TargetHeight, height);
                    } else media.UnknownPlacement = true;
                }
            }
        }
        return result;
    }

    private static bool TryGetRequiredPixels(OpenXmlElement reference, double dpi, out int width, out int height) {
        width = height = 0;
        double inchesWidth;
        double inchesHeight;
        double visibleWidth = 1, visibleHeight = 1;
        if (reference is A.Blip blip) {
            P.Picture? picture = blip.Ancestors<P.Picture>().FirstOrDefault();
            if (picture == null) return false;
            OpenXmlElement? fill = blip.Parent;
            if (fill?.ChildElements.Any(element => element.LocalName == "tile") == true) return false;
            // Non-default destination rectangles need their own transform mapping.
            if (fill?.Descendants<A.FillRectangle>().Any(rectangle => rectangle.GetAttributes().Any(attribute => attribute.Value != "0")) == true)
                return false;
            A.SourceRectangle? crop = fill?.GetFirstChild<A.SourceRectangle>();
            if (!TryVisibleFraction(crop?.Left?.Value ?? 0, crop?.Right?.Value ?? 0, out visibleWidth) ||
                !TryVisibleFraction(crop?.Top?.Value ?? 0, crop?.Bottom?.Value ?? 0, out visibleHeight)) return false;
            var drawing = picture.Ancestors<W.Drawing>().FirstOrDefault();
            if (drawing == null || drawing.Descendants().Any(element => element.LocalName == "sizeRelH" || element.LocalName == "sizeRelV")) return false;
            DW.Extent? outerExtent = drawing.Inline?.Extent ?? drawing.Anchor?.Extent;
            bool grouped = picture.Ancestors().Any(IsGroup);
            var extent = picture.ShapeProperties?.Transform2D?.Extents;
            long? cx = grouped ? extent?.Cx?.Value : outerExtent?.Cx?.Value;
            long? cy = grouped ? extent?.Cy?.Value : outerExtent?.Cy?.Value;
            if (!cx.HasValue || !cy.HasValue || cx <= 0 || cy <= 0) return false;
            inchesWidth = cx.Value / 914400D;
            inchesHeight = cy.Value / 914400D;
            // Use the maximum axis scale for nested groups, so rotation/shear cannot under-resolve either axis.
            foreach (OpenXmlElement group in picture.Ancestors().Where(IsGroup)) {
                OpenXmlElement? transform = group.ChildElements.FirstOrDefault(element => element.LocalName == "grpSpPr")?
                    .ChildElements.FirstOrDefault(element => element.LocalName == "xfrm");
                if (!TryGroupScale(transform, out double scale)) return false;
                inchesWidth *= scale; inchesHeight *= scale;
            }
            if (grouped) {
                OpenXmlElement topGroup = picture.Ancestors().Where(IsGroup).Last();
                OpenXmlElement? topExtent = topGroup.ChildElements.FirstOrDefault(element => element.LocalName == "grpSpPr")?
                    .ChildElements.FirstOrDefault(element => element.LocalName == "xfrm")?
                    .ChildElements.FirstOrDefault(element => element.LocalName == "ext");
                if (outerExtent?.Cx?.Value is not long outerX || outerExtent.Cy?.Value is not long outerY ||
                    !TryAttribute(topExtent, "cx", out double groupX) || !TryAttribute(topExtent, "cy", out double groupY) ||
                    outerX <= 0 || outerY <= 0 || groupX <= 0 || groupY <= 0) return false;
                double outerScale = Math.Max(outerX / groupX, outerY / groupY);
                inchesWidth *= outerScale; inchesHeight *= outerScale;
            }
        } else if (reference is V.ImageData vml) {
            V.Shape? shape = vml.Ancestors<V.Shape>().FirstOrDefault();
            if (shape == null || shape.Ancestors<V.Group>().Any()) return false;
            double? pixelsWidth = WordVmlImageGeometry.ReadDimension(shape.Style?.Value, "width");
            double? pixelsHeight = WordVmlImageGeometry.ReadDimension(shape.Style?.Value, "height");
            if (!pixelsWidth.HasValue || !pixelsHeight.HasValue) return false;
            if (!TryVmlCrop(vml, "cropleft", out double left) || !TryVmlCrop(vml, "cropright", out double right) ||
                !TryVmlCrop(vml, "croptop", out double top) || !TryVmlCrop(vml, "cropbottom", out double bottom)) return false;
            visibleWidth = 1 - left - right; visibleHeight = 1 - top - bottom;
            inchesWidth = pixelsWidth.Value / 96D; inchesHeight = pixelsHeight.Value / 96D;
        } else return false;
        return TryPixels(inchesWidth / visibleWidth, dpi, out width) &&
            TryPixels(inchesHeight / visibleHeight, dpi, out height);
    }

    private static bool IsGroup(OpenXmlElement element) => element.LocalName == "wgp" || element.LocalName == "grpSp";

    private static bool TryGroupScale(OpenXmlElement? transform, out double scale) {
        scale = 0;
        OpenXmlElement? extent = transform?.ChildElements.FirstOrDefault(element => element.LocalName == "ext");
        OpenXmlElement? childExtent = transform?.ChildElements.FirstOrDefault(element => element.LocalName == "chExt");
        if (!TryAttribute(extent, "cx", out double x) || !TryAttribute(extent, "cy", out double y) ||
            !TryAttribute(childExtent, "cx", out double childX) || !TryAttribute(childExtent, "cy", out double childY) ||
            x <= 0 || y <= 0 || childX <= 0 || childY <= 0) return false;
        scale = Math.Max(x / childX, y / childY);
        return !double.IsInfinity(scale) && !double.IsNaN(scale);
    }

    private static bool TryAttribute(OpenXmlElement? element, string name, out double value) =>
        double.TryParse(element?.GetAttributes().FirstOrDefault(attribute => attribute.LocalName == name).Value,
            NumberStyles.Float, CultureInfo.InvariantCulture, out value);

    private static bool TryVisibleFraction(int first, int second, out double fraction) {
        fraction = 1D - first / 100000D - second / 100000D;
        return first >= 0 && second >= 0 && fraction > 0 && fraction <= 1;
    }

    private static bool TryVmlCrop(V.ImageData image, string name, out double value) {
        string? text = image.GetAttributes().FirstOrDefault(attribute => attribute.LocalName == name).Value;
        value = 0;
        if (string.IsNullOrEmpty(text)) return true;
        bool fixedPoint = text!.EndsWith("f", StringComparison.OrdinalIgnoreCase);
        if (fixedPoint) text = text.Substring(0, text.Length - 1);
        if (!double.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out value)) return false;
        if (fixedPoint) value /= 65536D;
        return value >= 0 && value < 1;
    }

    private static bool TryPixels(double inches, double dpi, out int pixels) {
        pixels = 0;
        double required = Math.Ceiling(inches * dpi);
        if (double.IsNaN(required) || double.IsInfinity(required) || required <= 0 || required > int.MaxValue) return false;
        pixels = (int)required;
        return true;
    }
}
