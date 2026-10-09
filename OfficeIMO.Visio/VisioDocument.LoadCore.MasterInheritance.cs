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
        double scaleX = shape.Parent?.MasterShape?.Width > 0 ? shape.Parent.Width / shape.Parent.MasterShape.Width : 1;
        double scaleY = shape.Parent?.MasterShape?.Height > 0 ? shape.Parent.Height / shape.Parent.MasterShape.Height : 1;
        if (!shape.HasExplicitWidth) shape.Width = source.Width * scaleX;
        if (!shape.HasExplicitHeight) shape.Height = source.Height * scaleY;
        if (!shape.HasExplicitLocPinX) shape.LocPinX = source.LocPinX * scaleX;
        if (!shape.HasExplicitLocPinY) shape.LocPinY = source.LocPinY * scaleY;
        // Resolved zero pins are meaningful coordinates, not authoring defaults.
        shape.HasExplicitLocPinX = true; shape.HasExplicitLocPinY = true;
        if (!cells.Contains("PinX")) shape.PinX = source.PinX * scaleX;
        if (!cells.Contains("PinY")) shape.PinY = source.PinY * scaleY;
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
            shape.PreservedTextElement = source.PreservedTextElement == null ? null : new XElement(source.PreservedTextElement);
            shape.HasInheritedText = true;
        }
        if (source.TextStyle != null) {
            var style = source.TextStyle.Clone();
            style.ScaleTextBlock(source.Width > 0 ? shape.Width / source.Width : 1, source.Height > 0 ? shape.Height / source.Height : 1);
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
