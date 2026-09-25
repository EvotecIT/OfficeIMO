using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.OpenDocument;
using OfficeIMO.PowerPoint;
using System.Xml.Linq;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace OfficeIMO.PowerPoint.OpenDocument;

public static partial class PowerPointOpenDocumentConversionExtensions {
    private static int CountUnmappedPowerPointShapeAppearance(PresentationPart? presentation,
        IReadOnlyList<P.SlideId> slideIds) {
        if (presentation == null) return 0;
        int count = 0;
        foreach (P.SlideId slideId in slideIds) {
            if (slideId.RelationshipId?.Value is not string id ||
                presentation.GetPartById(id) is not SlidePart part) continue;
            if (part.Slide == null) continue;
            count += part.Slide.Descendants<P.ShapeProperties>().Count(HasUnmappedPowerPointShapeAppearance);
            count += part.Slide.Descendants<P.ShapeStyle>().Count();
            count += part.Slide.Descendants<P.Shape>().Count(shape =>
                shape.TextBody?.BodyProperties is A.BodyProperties body && (body.HasAttributes || body.HasChildren));
            count += part.Slide.Descendants<P.Picture>().Count(HasUnmappedPictureBlipAppearance);
        }
        return count;
    }

    private static int CountUnmappedPowerPointNonTextPlaceholders(PresentationPart? presentation,
        IReadOnlyList<P.SlideId> slideIds) {
        if (presentation == null) return 0;
        int count = 0;
        foreach (P.SlideId slideId in slideIds) {
            if (slideId.RelationshipId?.Value is not string id ||
                presentation.GetPartById(id) is not SlidePart part || part.Slide == null) continue;
            count += part.Slide.Descendants<P.Picture>().Count(picture =>
                picture.NonVisualPictureProperties?.ApplicationNonVisualDrawingProperties?
                    .GetFirstChild<P.PlaceholderShape>() != null);
            count += part.Slide.Descendants<P.GraphicFrame>().Count(frame =>
                frame.NonVisualGraphicFrameProperties?.ApplicationNonVisualDrawingProperties?
                    .GetFirstChild<P.PlaceholderShape>() != null);
            count += part.Slide.Descendants<P.ConnectionShape>().Count(connection =>
                connection.NonVisualConnectionShapeProperties?.ApplicationNonVisualDrawingProperties?
                    .GetFirstChild<P.PlaceholderShape>() != null);
            count += part.Slide.Descendants<P.Shape>().Count(shape =>
                shape.TextBody == null && shape.NonVisualShapeProperties?
                    .ApplicationNonVisualDrawingProperties?.GetFirstChild<P.PlaceholderShape>() != null);
        }
        return count;
    }

    private static int CountUnmappedPowerPointTextPlaceholderMetadata(PresentationPart? presentation,
        IReadOnlyList<P.SlideId> slideIds) {
        if (presentation == null) return 0;
        int count = 0;
        foreach (P.SlideId slideId in slideIds) {
            if (slideId.RelationshipId?.Value is not string id ||
                presentation.GetPartById(id) is not SlidePart part || part.Slide == null) continue;
            count += part.Slide.Descendants<P.Shape>().Count(shape =>
                shape.TextBody != null && shape.NonVisualShapeProperties?
                    .ApplicationNonVisualDrawingProperties?.GetFirstChild<P.PlaceholderShape>() is
                    P.PlaceholderShape placeholder &&
                (placeholder.Index != null || placeholder.Size != null || placeholder.Orientation != null));
        }
        return count;
    }

    private static int CountUnmappedPowerPointTextColors(PresentationPart? presentation,
        IReadOnlyList<P.SlideId> slideIds) {
        if (presentation == null) return 0;
        int count = 0;
        foreach (P.SlideId slideId in slideIds) {
            if (slideId.RelationshipId?.Value is not string id ||
                presentation.GetPartById(id) is not SlidePart part || part.Slide == null) continue;
            count += CountUnmappedPowerPointTextColors(part.Slide);
            if (part.NotesSlidePart?.NotesSlide is P.NotesSlide notes)
                count += CountUnmappedPowerPointTextColors(notes);
        }
        return count;
    }

    private static int CountUnmappedPowerPointTextColors(OpenXmlElement root) =>
        root.Descendants<A.RunProperties>().Count(HasUnmappedPowerPointTextColor) +
        root.Descendants<A.DefaultRunProperties>().Count(HasUnmappedPowerPointDefaultTextColor) +
        root.Descendants<A.EndParagraphRunProperties>().Count(HasUnmappedPowerPointDefaultTextColor);

    private static int CountUnmappedPowerPointTextTypography(PresentationPart? presentation,
        IReadOnlyList<P.SlideId> slideIds) {
        if (presentation == null) return 0;
        int count = 0;
        foreach (P.SlideId slideId in slideIds) {
            if (slideId.RelationshipId?.Value is not string id ||
                presentation.GetPartById(id) is not SlidePart part || part.Slide == null) continue;
            count += CountUnmappedPowerPointTextTypography(part.Slide);
            if (part.NotesSlidePart?.NotesSlide is P.NotesSlide notes)
                count += CountUnmappedPowerPointTextTypography(notes);
        }
        return count;
    }

    private static int CountUnmappedPowerPointTextTypography(OpenXmlElement root) =>
        root.Descendants<A.RunProperties>().Count(HasUnmappedPowerPointTextTypography) +
        root.Descendants<A.DefaultRunProperties>().Count(HasUnmappedPowerPointInheritedRunFormatting) +
        root.Descendants<A.EndParagraphRunProperties>().Count(HasUnmappedPowerPointInheritedRunFormatting);

    private static bool HasUnmappedPowerPointTextTypography(OpenXmlElement properties) =>
        properties.GetAttributes().Any(attribute => attribute.LocalName is "spc" or "kern");

    private static bool HasUnmappedPowerPointInheritedRunFormatting(OpenXmlElement properties) =>
        properties.GetAttributes().Any(attribute => attribute.LocalName is
            "b" or "i" or "sz" or "u" or "strike" or "baseline" or "cap" or "spc" or "kern") ||
        properties.HasChildren;

    private static bool HasUnmappedPowerPointTextColor(OpenXmlElement properties) {
        OpenXmlElement[] fills = properties.ChildElements.Where(IsFillElement).ToArray();
        return fills.Length > 1 || fills.Length == 1 && !IsDirectRgbFill(fills[0]) ||
            properties.Elements<A.Highlight>().Any(highlight =>
                highlight.ChildElements.Count != 1 ||
                highlight.GetFirstChild<A.RgbColorModelHex>() is not { ChildElements.Count: 0 });
    }

    private static bool HasUnmappedPowerPointDefaultTextColor(OpenXmlElement properties) =>
        properties.ChildElements.Any(child => IsFillElement(child) || child is A.Highlight);

    private static int CountUnmappedPowerPointShapeAccessibility(PresentationPart? presentation,
        IReadOnlyList<P.SlideId> slideIds) {
        if (presentation == null) return 0;
        int count = 0;
        foreach (P.SlideId slideId in slideIds) {
            if (slideId.RelationshipId?.Value is not string id ||
                presentation.GetPartById(id) is not SlidePart part || part.Slide == null) continue;
            count += CountUnmappedPowerPointShapeAccessibility(part.Slide);
            if (part.NotesSlidePart?.NotesSlide is P.NotesSlide notes)
                count += CountUnmappedPowerPointShapeAccessibility(notes);
        }
        return count;
    }

    private static int CountUnmappedPowerPointShapeAccessibility(OpenXmlElement root) =>
        root.Descendants<P.NonVisualDrawingProperties>().Count(properties =>
            properties.GetAttributes().Any(attribute => attribute.LocalName is "title" or "descr") ||
            properties.Descendants().Any(child => child.LocalName == "decorative"));

    private static bool HasUnmappedPowerPointShapeAppearance(P.ShapeProperties properties) {
        OpenXmlElement[] fills = properties.ChildElements.Where(IsFillElement).ToArray();
        if (fills.Length > 1 || fills.Length == 1 && !IsDirectRgbFill(fills[0]) &&
            !(fills[0] is A.NoFill && IsLineShape(properties))) return true;
        A.Outline? outline = properties.GetFirstChild<A.Outline>();
        if (outline != null && !IsDirectRgbOutline(outline)) return true;
        return properties.ChildElements.Any(child => child.LocalName is "effectLst" or "effectDag" or "scene3d" or "sp3d");
    }

    private static bool IsFillElement(OpenXmlElement element) => element is
        A.SolidFill or A.GradientFill or A.BlipFill or A.PatternFill or A.GroupFill or A.NoFill;

    private static bool IsLineShape(P.ShapeProperties properties) =>
        properties.Parent is P.ConnectionShape ||
        properties.GetFirstChild<A.PresetGeometry>()?.Preset?.Value == A.ShapeTypeValues.Line;

    private static bool IsDirectRgbFill(OpenXmlElement element) =>
        element is A.SolidFill solid && solid.ChildElements.Count == 1 &&
        solid.RgbColorModelHex is { ChildElements.Count: 0 };

    private static bool IsDirectRgbOutline(A.Outline outline) {
        if (outline.GetAttributes().Any(attribute => attribute.LocalName != "w")) return false;
        return outline.ChildElements.Count == 1 && IsDirectRgbFill(outline.ChildElements[0]);
    }

    private static bool HasUnmappedPictureBlipAppearance(P.Picture picture) {
        P.BlipFill? fill = picture.BlipFill;
        if (fill == null) return false;
        A.Blip? blip = fill.Blip;
        return blip?.ChildElements.Count > 0 ||
            fill.ChildElements.Any(child => child is not A.Blip and not A.SourceRectangle and not A.Stretch) ||
            fill.GetFirstChild<A.SourceRectangle>() is A.SourceRectangle crop &&
                crop.GetAttributes().Any(attribute => attribute.LocalName is "l" or "t" or "r" or "b" &&
                    (!int.TryParse(attribute.Value, out int value) || value < 0 || value > 100000)) ||
            fill.Descendants<A.FillRectangle>().Any(rectangle => rectangle.HasAttributes);
    }

    private static bool HasAuthoredPowerPointShapeLocks(OpenXmlElement element) =>
        element.LocalName is "spLocks" or "picLocks" or "cxnSpLocks" or "graphicFrameLocks" &&
        element.GetAttributes().Any(attribute => !IsStockPowerPointShapeLock(element.LocalName, attribute.LocalName) &&
            (attribute.Value == "1" || string.Equals(attribute.Value, "true", StringComparison.OrdinalIgnoreCase)));

    private static bool IsStockPowerPointShapeLock(string element, string attribute) =>
        element == "spLocks" && attribute == "noGrp" ||
        (element is "picLocks" or "graphicFrameLocks") && attribute == "noChangeAspect";

    private static int CountUnmappedPowerPointShapeLocks(PresentationPart? presentation,
        IReadOnlyList<P.SlideId> slideIds) {
        if (presentation == null) return 0;
        int count = 0;
        foreach (P.SlideId slideId in slideIds) {
            if (slideId.RelationshipId?.Value is not string id ||
                presentation.GetPartById(id) is not SlidePart part || part.Slide == null) continue;
            count += part.Slide.Descendants().Count(HasAuthoredPowerPointShapeLocks);
        }
        return count;
    }

    private static int CountUnmappedOdpTextLayout(OdpPresentation source) {
        XDocument content = source.Package.GetXml("content.xml");
        int paragraphs = content.Descendants().Count(paragraph =>
            (paragraph.Name == OdfNamespaces.Text + "p" || paragraph.Name == OdfNamespaces.Text + "h") &&
            HasUnmappedOdpTextLayout(source, (string?)paragraph.Attribute(OdfNamespaces.Text + "style-name")));
        int inlineStyles = content.Descendants().Count(element =>
            (element.Name == OdfNamespaces.Text + "span" || element.Name == OdfNamespaces.Text + "a") &&
            HasUnmappedOdpCharacterSpacing(source, OdfStyleFamily.Text,
                (string?)element.Attribute(OdfNamespaces.Text + "style-name")));
        return paragraphs + inlineStyles;
    }

    private static bool HasUnmappedOdpTextLayout(OdpPresentation source, string? styleName) {
        XElement? paragraph = EffectiveOdfStyleProperties(source, OdfStyleFamily.Paragraph, styleName,
            OdfNamespaces.Style + "paragraph-properties");
        XElement? authored = EffectiveOdfStyleProperties(source, OdfStyleFamily.Paragraph, styleName,
            OdfNamespaces.Style + "paragraph-properties", includeDefault: false);
        if (paragraph?.Attributes().Any(attribute => authored?.Attribute(attribute.Name) == null) == true)
            return true;
        if (paragraph != null && (paragraph.HasElements || paragraph.Attributes().Any(attribute =>
            attribute.Name != OdfNamespaces.Fo + "text-align" &&
            attribute.Name != OdfNamespaces.Fo + "line-height" &&
            attribute.Name != OdfNamespaces.Style + "writing-mode"))) return true;
        return HasUnmappedOdpCharacterSpacing(source, OdfStyleFamily.Paragraph, styleName);
    }

    private static bool HasUnmappedOdpCharacterSpacing(OdpPresentation source, OdfStyleFamily family,
        string? styleName) {
        XElement? text = EffectiveOdfStyleProperties(source, family, styleName,
            OdfNamespaces.Style + "text-properties");
        return text?.Attributes().Any(attribute =>
            attribute.Name == OdfNamespaces.Fo + "letter-spacing" ||
            attribute.Name == OdfNamespaces.Fo + "word-spacing" ||
            attribute.Name == OdfNamespaces.Fo + "text-shadow" ||
            attribute.Name == OdfNamespaces.Style + "text-rotation-angle") == true;
    }

    private static int CountUnmappedOdpTextEffects(OdpPresentation source) {
        XDocument content = source.Package.GetXml("content.xml");
        return content.Descendants().Count(element => {
            OdfStyleFamily family;
            if (element.Name == OdfNamespaces.Text + "p" || element.Name == OdfNamespaces.Text + "h")
                family = OdfStyleFamily.Paragraph;
            else if (element.Name == OdfNamespaces.Text + "span" || element.Name == OdfNamespaces.Text + "a")
                family = OdfStyleFamily.Text;
            else return false;
            XElement? properties = EffectiveOdfStyleProperties(source, family,
                (string?)element.Attribute(OdfNamespaces.Text + "style-name"),
                OdfNamespaces.Style + "text-properties");
            XElement? authored = EffectiveOdfStyleProperties(source, family,
                (string?)element.Attribute(OdfNamespaces.Text + "style-name"),
                OdfNamespaces.Style + "text-properties", includeDefault: false);
            if (properties?.Attributes().Any(attribute =>
                authored?.Attribute(attribute.Name) == null && !IsEquivalentOdpDefaultTextProperty(attribute)) == true)
                return true;
            return properties != null && (properties.HasElements ||
                properties.Attributes().Any(attribute => !IsMappedOdpTextProperty(attribute)));
        });
    }

    private static bool IsEquivalentOdpDefaultTextProperty(XAttribute attribute) =>
        (attribute.Name == OdfNamespaces.Fo + "font-weight" ||
         attribute.Name == OdfNamespaces.Fo + "font-style") &&
        attribute.Value == "normal";

    private static bool IsMappedOdpTextProperty(XAttribute attribute) {
        if (attribute.IsNamespaceDeclaration) return true;
        XName name = attribute.Name;
        if (name == OdfNamespaces.Fo + "font-weight")
            return attribute.Value is "bold" or "normal";
        if (name == OdfNamespaces.Fo + "font-style")
            return attribute.Value is "italic" or "normal";
        return name == OdfNamespaces.Fo + "font-size" ||
            name == OdfNamespaces.Fo + "font-family" ||
            name == OdfNamespaces.Style + "font-name" ||
            name == OdfNamespaces.Fo + "color" ||
            name == OdfNamespaces.Fo + "background-color" ||
            name == OdfNamespaces.Style + "text-underline-style" ||
            name == OdfNamespaces.Style + "text-underline-type" ||
            name == OdfNamespaces.Style + "text-line-through-style" ||
            name == OdfNamespaces.Style + "text-line-through-type" ||
            name == OdfNamespaces.Style + "text-position" ||
            name == OdfNamespaces.Fo + "text-transform" ||
            name == OdfNamespaces.Fo + "font-variant" ||
            // These have their own paragraph-layout loss count.
            name == OdfNamespaces.Fo + "letter-spacing" ||
            name == OdfNamespaces.Fo + "word-spacing" ||
            name == OdfNamespaces.Fo + "text-shadow" ||
            name == OdfNamespaces.Style + "text-rotation-angle";
    }

    private static int CountUnmappedOdpTableProtection(OdpPresentation source) {
        int count = 0;
        foreach (OdpSlide slide in source.Slides) {
            foreach (OdpTable table in slide.Shapes.OfType<OdpTable>()) {
                if (table.Element.DescendantsAndSelf().Any(element => element.Attributes().Any(attribute =>
                    attribute.Name.Namespace == OdfNamespaces.Table &&
                    ((attribute.Name.LocalName == "protected" && (attribute.Value is "true" or "1")) ||
                     attribute.Name.LocalName.StartsWith("protection-", StringComparison.Ordinal)))))
                    count++;
            }
        }
        return count;
    }

    private static int CountUnmappedOdpShapeLayers(OdpPresentation source) =>
        source.Slides.Sum(slide => slide.Element.Descendants().Count(element =>
            element.Attribute(OdfNamespaces.Draw + "layer") != null));

    private static int CountUnmappedOdpNavigationOrder(OdpPresentation source) =>
        source.Slides.Count(slide => slide.Element.Attribute(OdfNamespaces.Draw + "nav-order") != null);

    private static int CountUnmappedPowerPointTextGeometry(PresentationPart? presentation,
        IReadOnlyList<P.SlideId> slideIds) {
        if (presentation == null) return 0;
        int count = 0;
        foreach (P.SlideId slideId in slideIds) {
            if (slideId.RelationshipId?.Value is not string id ||
                presentation.GetPartById(id) is not SlidePart part) continue;
            count += part.Slide?.Descendants<P.Shape>().Count(shape => {
                if (shape.TextBody == null) return false;
                P.ShapeProperties? properties = shape.ShapeProperties;
                A.PresetGeometry? preset = properties?.GetFirstChild<A.PresetGeometry>();
                return properties?.GetFirstChild<A.CustomGeometry>() != null ||
                    preset != null && (preset.Preset?.Value != A.ShapeTypeValues.Rectangle ||
                                       preset.AdjustValueList?.HasChildren == true);
            }) ?? 0;
        }
        return count;
    }

    private static int CountUnmappedPowerPointParagraphLayout(PresentationPart? presentation,
        IReadOnlyList<P.SlideId> slideIds) {
        if (presentation == null) return 0;
        int count = 0;
        foreach (P.SlideId slideId in slideIds) {
            if (slideId.RelationshipId?.Value is not string id ||
                presentation.GetPartById(id) is not SlidePart part || part.Slide == null) continue;
            count += CountUnmappedPowerPointParagraphLayout(part.Slide);
            if (part.NotesSlidePart?.NotesSlide is P.NotesSlide notes)
                count += CountUnmappedPowerPointParagraphLayout(notes);
        }
        return count;
    }

    private static int CountUnmappedPowerPointParagraphLayout(OpenXmlElement root) =>
        root.Descendants<A.ParagraphProperties>().Count(properties =>
                properties.GetAttributes().Any(attribute =>
                    attribute.LocalName is not "algn" and not "rtl" &&
                    (attribute.LocalName != "lvl" || attribute.Value != "0")) ||
                properties.ChildElements.Any(child => child.LocalName is "spcBef" or "spcAft" or "tabLst")) +
        root.Descendants<A.ListStyle>().Sum(style => style.ChildElements.Count(level =>
                level.LocalName.StartsWith("lvl", StringComparison.Ordinal) &&
                level.LocalName.EndsWith("pPr", StringComparison.Ordinal) &&
                (level.HasAttributes || level.HasChildren)));

    private static int CountUnmappedPowerPointTransitionTiming(PresentationPart? presentation,
        IReadOnlyList<P.SlideId> slideIds) {
        if (presentation == null) return 0;
        int count = 0;
        foreach (P.SlideId slideId in slideIds) {
            if (slideId.RelationshipId?.Value is not string id ||
                presentation.GetPartById(id) is not SlidePart part) continue;
            if (part.Slide?.Transition is P.Transition transition &&
                (transition.GetAttributes().Any(attribute =>
                    attribute.LocalName is "advTm" or "advClick" or "spd" or "dur") ||
                 transition.GetFirstChild<P.SoundAction>() != null)) count++;
        }
        return count;
    }

    private static bool HasPowerPointDynamicSlideBackground(PresentationPart? presentation, P.SlideId slideId) {
        if (presentation == null || slideId.RelationshipId?.Value is not string id ||
            presentation.GetPartById(id) is not SlidePart part) return false;
        P.Background? background = part.Slide?.CommonSlideData?.Background;
        return background?.ChildElements.Any(child => child.LocalName == "bgRef") == true ||
            background?.Descendants().Any(child => child.LocalName is "schemeClr" or "sysClr" or "effectLst" or "effectDag") == true ||
            background?.BackgroundProperties?.GetAttributes().Any(attribute =>
                attribute.LocalName == "shadeToTitle" && attribute.Value is "1" or "true") == true;
    }

    private static int CountUnwrappedOdpDrawingElements(OdpSlide slide) {
        var wrapped = new HashSet<XElement>(slide.Shapes.Select(shape => shape.Element));
        return slide.Element.Elements().Count(element =>
            element.Name.Namespace == OdfNamespaces.Draw &&
            element.Name != OdfNamespaces.Draw + "page-thumbnail" && !wrapped.Contains(element));
    }

    private static bool HasUnmappedOdpShapeAppearance(OdpPresentation source, OdpShape shape) {
        if (shape.Element.Attribute(OdfNamespaces.Presentation + "style-name") != null ||
            HasUnmappedOdpDirectShapeGeometry(shape)) return true;
        string? styleName = (string?)shape.Element.Attribute(OdfNamespaces.Draw + "style-name");
        XElement? properties = EffectiveOdfStyleProperties(source, OdfStyleFamily.Graphic, styleName,
            OdfNamespaces.Style + "graphic-properties");
        if (properties == null) return false;
        string? fill = (string?)properties.Attribute(OdfNamespaces.Draw + "fill");
        string? stroke = (string?)properties.Attribute(OdfNamespaces.Draw + "stroke");
        if (fill != null && fill != "solid" && !(shape is OdpLine && fill == "none") ||
            stroke != null && stroke != "solid") return true;
        if (fill == "solid" && !OdfColor.TryParse((string?)properties.Attribute(OdfNamespaces.Draw + "fill-color"), out _)) return true;
        if (stroke == "solid" && !OdfColor.TryParse((string?)properties.Attribute(OdfNamespaces.Svg + "stroke-color"), out _)) return true;
        bool mappedImageClip = shape is OdpImage image && image.Crop.HasValue;
        return properties.HasElements || properties.Attributes().Any(attribute =>
            attribute.Name != OdfNamespaces.Draw + "fill" && attribute.Name != OdfNamespaces.Draw + "fill-color" &&
            attribute.Name != OdfNamespaces.Draw + "stroke" && attribute.Name != OdfNamespaces.Svg + "stroke-color" &&
            attribute.Name != OdfNamespaces.Svg + "stroke-width" &&
            (attribute.Name != OdfNamespaces.Fo + "clip" || !mappedImageClip));
    }

    private static bool HasUnmappedOdpDirectShapeGeometry(OdpShape shape) {
        if (shape is not OdpRectangle and not OdpEllipse and not OdpLine) return false;
        return shape.Element.Attributes().Any(attribute =>
            !attribute.IsNamespaceDeclaration &&
            attribute.Name != OdfNamespaces.Draw + "name" &&
            attribute.Name != OdfNamespaces.Draw + "style-name" &&
            attribute.Name != OdfNamespaces.Draw + "transform" &&
            attribute.Name != OdfNamespaces.Presentation + "visibility" &&
            attribute.Name != OdfNamespaces.Presentation + "class" &&
            attribute.Name != XNamespace.Xml + "id" &&
            attribute.Name != OdfNamespaces.Svg + "x" &&
            attribute.Name != OdfNamespaces.Svg + "y" &&
            attribute.Name != OdfNamespaces.Svg + "width" &&
            attribute.Name != OdfNamespaces.Svg + "height" &&
            attribute.Name != OdfNamespaces.Svg + "x1" &&
            attribute.Name != OdfNamespaces.Svg + "y1" &&
            attribute.Name != OdfNamespaces.Svg + "x2" &&
            attribute.Name != OdfNamespaces.Svg + "y2");
    }

    private static bool HasUnmappedOdpNoteContent(OdpPresentation source, OdpSlide slide) {
        XElement? notes = slide.Element.Element(OdfNamespaces.Presentation + "notes");
        if (notes == null) return false;
        bool defaultFrameAppearance = notes.Descendants(OdfNamespaces.Draw + "frame").Any() &&
            source.Package.GetXml("styles.xml").Descendants(OdfNamespaces.Style + "default-style")
                .Any(candidate => (string?)candidate.Attribute(OdfNamespaces.Style + "family") == "graphic" &&
                    candidate.Element(OdfNamespaces.Style + "graphic-properties") is XElement properties &&
                    (properties.HasAttributes || properties.HasElements));
        return notes.HasAttributes || defaultFrameAppearance || notes.Descendants().Any(element =>
            element.Name == OdfNamespaces.Text + "h" ||
            element.Name == OdfNamespaces.Text + "list" ||
            element.Name.Namespace == OdfNamespaces.Table ||
            element.Name.Namespace == OdfNamespaces.Draw &&
            element.Name != OdfNamespaces.Draw + "frame" &&
            element.Name != OdfNamespaces.Draw + "text-box" &&
            element.Name != OdfNamespaces.Draw + "page-thumbnail" ||
            element.Name == OdfNamespaces.Draw + "frame" && element.Attributes().Any(attribute =>
                attribute.Name != OdfNamespaces.Draw + "name" &&
                (attribute.Name != OdfNamespaces.Presentation + "class" || attribute.Value != "notes")) ||
            element.Name == OdfNamespaces.Draw + "text-box" && element.HasAttributes);
    }

    private static bool HasUnmappedOdpTransitionTiming(OdpPresentation source, OdpSlide slide) {
        string? styleName = (string?)slide.Element.Attribute(OdfNamespaces.Draw + "style-name");
        XElement? properties = EffectiveOdfStyleProperties(source, OdfStyleFamily.DrawingPage, styleName,
            OdfNamespaces.Style + "drawing-page-properties");
        return properties?.Element(OdfNamespaces.Presentation + "sound") != null ||
            slide.Element.Element(OdfNamespaces.Presentation + "sound") != null ||
            properties?.Attribute(OdfNamespaces.Presentation + "transition-change") != null ||
            properties?.Attribute(OdfNamespaces.Presentation + "duration") != null ||
            slide.Element.Attribute(OdfNamespaces.Presentation + "transition-change") != null ||
            slide.Element.Attribute(OdfNamespaces.Presentation + "duration") != null;
    }

    private static int CountUnmappedPowerPointTableAppearance(PresentationPart? presentation,
        IReadOnlyList<P.SlideId> slideIds) {
        if (presentation == null) return 0;
        int count = 0;
        foreach (P.SlideId slideId in slideIds) {
            if (slideId.RelationshipId?.Value is not string id ||
                presentation.GetPartById(id) is not SlidePart part) continue;
            if (part.Slide != null)
                count += part.Slide.Descendants<A.Table>().Count(HasUnmappedPowerPointTableAppearance);
        }
        return count;
    }

    private static bool HasUnmappedPowerPointTableAppearance(A.Table table) {
        A.TableProperties? properties = table.TableProperties;
        if (properties != null && (properties.HasAttributes || properties.HasChildren)) return true;
        long[] columnWidths = table.TableGrid?.Elements<A.GridColumn>()
            .Select(column => column.Width?.Value ?? 0L).ToArray() ?? Array.Empty<long>();
        long[] rowHeights = table.Elements<A.TableRow>()
            .Select(row => row.Height?.Value ?? 0L).ToArray();
        if (columnWidths.Distinct().Skip(1).Any() || rowHeights.Distinct().Skip(1).Any()) return true;
        A.Extents? extents = table.Ancestors<P.GraphicFrame>().FirstOrDefault()?.Transform?.Extents;
        if (extents != null && (columnWidths.Sum() != extents.Cx?.Value ||
            rowHeights.Sum() != extents.Cy?.Value)) return true;
        return table.Descendants<A.TableCellProperties>().Any(cell => cell.HasAttributes || cell.HasChildren) ||
            table.Descendants<A.BodyProperties>().Any(body => body.HasAttributes || body.HasChildren);
    }

    private static bool HasUnmappedOdpTableAppearance(OdpTable table) =>
        table.Element.DescendantsAndSelf().Any(element => element.Attributes().Any(attribute =>
            attribute.Name == OdfNamespaces.Table + "style-name" ||
                attribute.Name == OdfNamespaces.Table + "default-cell-style-name" ||
                attribute.Name == OdfNamespaces.Table + "template-name" ||
                attribute.Name.Namespace == OdfNamespaces.Table &&
                attribute.Name.LocalName.StartsWith("use-", StringComparison.Ordinal) &&
                attribute.Name.LocalName.EndsWith("-styles", StringComparison.Ordinal) &&
                attribute.Value is "true" or "1"));

    private static bool HasUnmappedOdpShapeAccessibility(OdpShape shape) =>
        shape.Element.DescendantsAndSelf().Any(element =>
            element.Name == OdfNamespaces.Svg + "title" || element.Name == OdfNamespaces.Svg + "desc");

    private static bool HasUnmappedOdpTableVisibility(OdpTable table) =>
        table.Element.Descendants().Any(element =>
            (element.Name == OdfNamespaces.Table + "table-row" ||
             element.Name == OdfNamespaces.Table + "table-column") &&
            (string?)element.Attribute(OdfNamespaces.Table + "visibility") is "collapse" or "filter");

    private static bool HasUnmappedOdpTableLists(OdpTable table) =>
        table.Element.Descendants(OdfNamespaces.Table + "table-cell").Any(cell =>
            cell.Descendants(OdfNamespaces.Text + "list").Any());

    private static int CountUnmappedOdpEmbeddedFonts(OdpPresentation source) =>
        new[] { "content.xml", "styles.xml" }.Sum(part => source.Package.GetXml(part)
            .Descendants(OdfNamespaces.Style + "font-face")
            .Count(face => face.Descendants(OdfNamespaces.Svg + "font-face-src").Any() ||
                           face.Descendants(OdfNamespaces.Svg + "font-face-uri").Any()));

    private static bool HasUnmappedOdpTableValues(OdpTable table) =>
        table.Element.Descendants(OdfNamespaces.Table + "table-cell").Any(cell =>
            cell.Attributes().Any(attribute => attribute.Name.Namespace == OdfNamespaces.Office &&
                attribute.Name.LocalName is "value-type" or "value" or "date-value" or "time-value" or
                    "boolean-value" or "string-value" or "currency"));

    private static int CountUnmappedOdpCustomShows(OdpPresentation source) =>
        source.Package.GetXml("content.xml").Descendants(OdfNamespaces.Presentation + "show").Count();

    private static int CountUnmappedOdpSlideShowSettings(OdpPresentation source) =>
        source.Package.GetXml("content.xml").Descendants(OdfNamespaces.Presentation + "settings")
            .Count(settings => settings.HasAttributes || settings.Elements()
                .Any(element => element.Name != OdfNamespaces.Presentation + "show"));

    private static (bool Override, bool Loss, OdfColor? Color, bool SuppressesMasterBackground) ReadOdpSlideBackground(
        OdpPresentation source, OdpSlide slide) {
        string? styleName = (string?)slide.Element.Attribute(OdfNamespaces.Draw + "style-name");
        XElement? authoredProperties = EffectiveOdfStyleProperties(source, OdfStyleFamily.DrawingPage, styleName,
            OdfNamespaces.Style + "drawing-page-properties", includeDefault: false);
        XElement? properties = EffectiveOdfStyleProperties(source, OdfStyleFamily.DrawingPage, styleName,
            OdfNamespaces.Style + "drawing-page-properties");
        bool suppressesMasterBackground = string.Equals(
            (string?)slide.Element.Attribute(OdfNamespaces.Presentation + "background-visible"),
            "false", StringComparison.OrdinalIgnoreCase) || string.Equals(
            (string?)properties?.Attribute(OdfNamespaces.Presentation + "background-visible"),
            "false", StringComparison.OrdinalIgnoreCase);
        if (properties == null) return (false, false, null, suppressesMasterBackground);
        string? fill = (string?)properties.Attribute(OdfNamespaces.Draw + "fill");
        if (fill == "none" && authoredProperties?.Attribute(OdfNamespaces.Draw + "fill") == null)
            fill = null;
        if (fill == null) {
            bool unsupportedInheritedProperties = properties.Attributes().Any(attribute =>
                attribute.Name.Namespace == OdfNamespaces.Draw &&
                attribute.Name != OdfNamespaces.Draw + "fill" &&
                attribute.Name != OdfNamespaces.Draw + "fill-color");
            return (false, unsupportedInheritedProperties, null, suppressesMasterBackground);
        }
        OdfColor? color = fill == "solid" &&
            OdfColor.TryParse((string?)properties.Attribute(OdfNamespaces.Draw + "fill-color"), out OdfColor parsed)
            ? parsed : (OdfColor?)null;
        bool loss = fill != "none" && !color.HasValue || properties.Attributes().Any(attribute =>
            attribute.Name.Namespace == OdfNamespaces.Draw &&
            attribute.Name != OdfNamespaces.Draw + "fill" && attribute.Name != OdfNamespaces.Draw + "fill-color");
        return (true, loss, color, suppressesMasterBackground);
    }

    private static XElement? EffectiveOdfStyleProperties(OdpPresentation source, OdfStyleFamily family,
        string? styleName, XName propertiesName, string partPath = "content.xml", bool includeDefault = true) {
        var effective = new XElement(propertiesName);
        OdfStyle? style = string.IsNullOrWhiteSpace(styleName) ? null :
            source.Styles.FindInPart(family, styleName!, partPath);
        foreach (OdfStyle candidate in style == null ? Array.Empty<OdfStyle>() : source.Styles.Resolve(style)) {
            XElement? properties = candidate.Element.Element(propertiesName);
            if (properties == null) continue;
            foreach (XAttribute attribute in properties.Attributes()) {
                if (effective.Attribute(attribute.Name) == null)
                    effective.SetAttributeValue(attribute.Name, attribute.Value);
            }
            foreach (XElement child in properties.Elements()) effective.Add(new XElement(child));
        }
        XElement? defaultProperties = includeDefault
            ? source.Styles.FindDefaultProperties(family, propertiesName) : null;
        if (defaultProperties != null) {
            foreach (XAttribute attribute in defaultProperties.Attributes()) {
                if (effective.Attribute(attribute.Name) == null)
                    effective.SetAttributeValue(attribute.Name, attribute.Value);
            }
            foreach (XElement child in defaultProperties.Elements()) {
                if (!effective.Elements(child.Name).Any()) effective.Add(new XElement(child));
            }
        }
        if ((string?)effective.Attribute(OdfNamespaces.Draw + "fill") == "solid") {
            effective.SetAttributeValue(OdfNamespaces.Draw + "fill-gradient-name", null);
            effective.SetAttributeValue(OdfNamespaces.Draw + "fill-image-name", null);
            effective.SetAttributeValue(OdfNamespaces.Draw + "fill-hatch-name", null);
            effective.SetAttributeValue(OdfNamespaces.Draw + "fill-transparency-gradient-name", null);
        }
        return effective.HasAttributes || effective.HasElements ? effective : null;
    }
}
