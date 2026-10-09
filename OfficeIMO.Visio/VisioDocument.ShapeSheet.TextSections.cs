using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;

namespace OfficeIMO.Visio;

public partial class VisioDocument {
    /// <summary>Synchronizes a ShapeSheet text-row assignment with its supported typed fields.</summary>
    internal static bool ModelShapeSheetTextSection(VisioShape shape, VisioShapeSheetSection assigned, bool character) {
        XElement section = assigned.ToXElement();
        VisioTextStyle? pendingFont = character ? PendingShapeSheetFontAssignment(assigned,
            shape.CharacterSectionSource, shape.TextStyle) : null;
        IReadOnlyDictionary<int, string> fonts = ShapeSheetFontNames(shape.TextStyle,
            shape.NativeFontScope?.FaceNames.Elements() ?? shape.OwnerPage?.OwnerDocument?.PreservedFaceNamesElements);
        bool modeled = character
            ? TryParseSimpleCharSection(shape, section, section.Name.Namespace, fonts)
            : TryParseSimpleParaSection(shape, section, section.Name.Namespace);
        if (modeled) RestorePendingFontAssignment(shape.TextStyle, pendingFont);
        if (character) {
            shape.HasModeledCharSection = modeled;
            shape.CharacterSectionSource = modeled ? CaptureTextSection(section, shape.TextStyle, true) : null;
        } else {
            shape.HasModeledParaSection = modeled;
            shape.ParagraphSectionSource = modeled ? CaptureTextSection(section, shape.TextStyle, false) : null;
        }
        if (!modeled) ClearShapeSheetTextFormatting(shape.TextStyle, character);
        return modeled;
    }

    /// <summary>Synchronizes a connector's ShapeSheet row without inventing font-table identities.</summary>
    internal static bool ModelShapeSheetTextSection(VisioConnector connector, VisioShapeSheetSection assigned, bool character) {
        XElement section = assigned.ToXElement();
        VisioTextStyle? pendingFont = character ? PendingShapeSheetFontAssignment(assigned,
            connector.CharacterSectionSource, connector.TextStyle) : null;
        // Connectors have no detached master font scope. A retained Font cell can
        // use its loaded identity; an unknown replacement remains native-owned.
        bool modeled = character
            ? TryParseSimpleConnectorCharSection(connector, section, section.Name.Namespace,
                ShapeSheetFontNames(connector.TextStyle, null))
            : TryParseSimpleConnectorParaSection(connector, section, section.Name.Namespace);
        if (modeled) RestorePendingFontAssignment(connector.TextStyle, pendingFont);
        if (character) {
            connector.HasModeledCharSection = modeled;
            connector.CharacterSectionSource = modeled ? CaptureTextSection(section, connector.TextStyle, true) : null;
        } else {
            connector.HasModeledParaSection = modeled;
            connector.ParagraphSectionSource = modeled ? CaptureTextSection(section, connector.TextStyle, false) : null;
        }
        if (!modeled) ClearShapeSheetTextFormatting(connector.TextStyle, character);
        return modeled;
    }

    private static VisioTextStyle? PendingShapeSheetFontAssignment(VisioShapeSheetSection assigned,
        VisioTextSectionSource? source, VisioTextStyle? style) {
        if (style?.FontFamilyAssigned != true || assigned.Rows.SelectMany(row => row.Cells)
                .Any(cell => cell.Name == "Font" && cell.ValueAssigned)) return null;
        XElement? current = GetRenderTextSection(Enumerable.Empty<XElement>(), source, style, character: true);
        XElement? Font(XElement? value) => value?.Elements(value.Name.Namespace + "Row")
            .Elements(value.Name.Namespace + "Cell").FirstOrDefault(cell => (string?)cell.Attribute("N") == "Font");
        return XNode.DeepEquals(Font(current), Font(assigned.ToXElement())) ? style.Clone() : null;
    }

    private static void RestorePendingFontAssignment(VisioTextStyle? style, VisioTextStyle? pending) {
        if (style == null || pending == null) return;
        style.FontFamily = pending.FontFamily;
        style.FontFaceId = pending.FontFaceId;
        style.FontFamilyAssigned = true;
    }

    private static IReadOnlyDictionary<int, string> ShapeSheetFontNames(VisioTextStyle? style,
        IEnumerable<XElement>? faces) {
        var names = new Dictionary<int, string>();
        foreach (XElement face in faces ?? Enumerable.Empty<XElement>()) {
            if (TryParseCellIntValue((string?)face.Attribute("ID"), out int id) &&
                (string?)face.Attribute("Name") is string name && !string.IsNullOrWhiteSpace(name))
                names[id] = name;
        }
        if (style?.FontFaceId is int known && !names.ContainsKey(known) && !string.IsNullOrWhiteSpace(style.FontFamily))
            names.Add(known, style.FontFamily!);
        return names;
    }

    /// <summary>Clears only the typed fields owned by a removed or complex native text section.</summary>
    internal static void ClearShapeSheetTextFormatting(VisioTextStyle? style, bool character) {
        if (style == null) return;
        if (!character) { style.HorizontalAlignment = null; return; }
        style.FontFaceId = null;
        style.FontFamily = null;
        style.FontFamilyAssigned = false;
        style.Color = null;
        style.Size = null;
        style.Bold = null;
        style.Italic = null;
        style.UnderlineStyle = null;
        style.StrikethroughStyle = null;
        style.SmallCaps = null;
        style.Capitalization = null;
        style.Baseline = null;
    }
}
