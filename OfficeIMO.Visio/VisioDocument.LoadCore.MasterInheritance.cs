using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;

namespace OfficeIMO.Visio;

public partial class VisioDocument {
    private static void ApplyInheritedMasterProperties(VisioShape shape, XElement element, XNamespace ns, VisioShape source) {
        var cells = new HashSet<string>(element.Elements(ns + "Cell").Attributes("N").Select(attribute => attribute.Value), StringComparer.Ordinal);
        if (element.Element(ns + "XForm") is XElement xform)
            foreach (XElement cell in xform.Elements()) cells.Add(cell.Name.LocalName);
        double scaleX = shape.Parent?.MasterShape?.WidthInches > 0 ? shape.Parent.WidthInches / shape.Parent.MasterShape.WidthInches : 1;
        double scaleY = shape.Parent?.MasterShape?.HeightInches > 0 ? shape.Parent.HeightInches / shape.Parent.MasterShape.HeightInches : 1;
        if (!shape.HasExplicitWidth) shape.WidthInches = source.WidthInches * scaleX;
        if (!shape.HasExplicitHeight) shape.HeightInches = source.HeightInches * scaleY;
        if (!shape.HasExplicitLocPinX) shape.LocPinXInches = source.LocPinXInches * scaleX;
        if (!shape.HasExplicitLocPinY) shape.LocPinYInches = source.LocPinYInches * scaleY;
        // Resolved zero pins are meaningful coordinates, not authoring defaults.
        shape.HasExplicitLocPinX = true; shape.HasExplicitLocPinY = true;
        if (!cells.Contains("PinX")) shape.PinXInches = source.PinXInches * scaleX;
        if (!cells.Contains("PinY")) shape.PinYInches = source.PinYInches * scaleY;
        if (!cells.Contains("Angle")) shape.Angle = source.Angle;
        if (string.IsNullOrEmpty(shape.NativeStyleReferences?.LineStyle)) {
            if (!cells.Contains("LineWeight")) shape.LineWeight = source.LineWeight;
            if (!cells.Contains("LineColor")) shape.LineColor = source.LineColor;
            else if (!cells.Contains("LineColorTrans")) shape.LineColor = VisioNativePaintStyleResolver.InheritTransparency(shape.LineColor, source.LineColor);
            if (!cells.Contains("LinePattern")) shape.LinePattern = source.LinePattern;
        }
        if (string.IsNullOrEmpty(shape.NativeStyleReferences?.FillStyle)) {
            if (!cells.Contains("FillForegnd")) shape.FillColor = source.FillColor;
            else if (!cells.Contains("FillForegndTrans")) shape.FillColor = VisioNativePaintStyleResolver.InheritTransparency(shape.FillColor, source.FillColor);
            if (!cells.Contains("FillPattern")) shape.FillPattern = source.FillPattern;
        }
        VisioNativePaintStyleResolver.ApplyLocalTransparency(shape, element);
        shape.Type ??= source.Type;
        if (element.Element(ns + "Text") == null) {
            shape.Text = source.Text;
            shape.PreservedTextValue = source.PreservedTextValue;
            // Inherited text is immutable during loading and projection. Share its XML
            // so a compact master cannot expand into a full tree for every instance.
            shape.PreservedTextElement = source.PreservedTextElement;
            shape.HasInheritedText = true;
        }
        if (source.TextStyle != null) {
            var style = source.TextStyle.Clone();
            style.ScaleTextBlock(source.WidthInches > 0 ? shape.WidthInches / source.WidthInches : 1, source.HeightInches > 0 ? shape.HeightInches / source.HeightInches : 1);
            bool preservedCharacter = shape.PreservedNonGeometrySections.Any(section => IsCharacterSection((string?)section.Attribute("N")));
            bool preservedParagraph = shape.PreservedNonGeometrySections.Any(section => IsParagraphSection((string?)section.Attribute("N")));
            bool explicitTextStyle = !string.IsNullOrEmpty(shape.NativeStyleReferences?.TextStyle);
            shape.TextStyle ??= new VisioTextStyle();
            shape.TextStyle.InheritUnsetFrom(style, inheritCharacter: !preservedCharacter && !explicitTextStyle,
                inheritParagraph: !preservedParagraph && !explicitTextStyle);
            if (!preservedCharacter && !explicitTextStyle)
                shape.CharacterSectionSource = CaptureInheritedTextSection(shape.CharacterSectionSource, source.CharacterSectionSource, shape.TextStyle, ns, character: true);
            if (!preservedParagraph && !explicitTextStyle)
                shape.ParagraphSectionSource = CaptureInheritedTextSection(shape.ParagraphSectionSource, source.ParagraphSectionSource, shape.TextStyle, ns, character: false);
        }
    }
}
