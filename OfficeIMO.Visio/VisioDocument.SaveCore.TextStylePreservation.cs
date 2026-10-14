using System.Linq;
using System.Xml;
using System.Xml.Linq;

namespace OfficeIMO.Visio;

public partial class VisioDocument {
    // Use the same edit merge for rendering and writing. Loaded model values can
    // be inherited caches, so only changes from the captured baseline are local.
    internal static XElement? GetRenderTextSection(IEnumerable<XElement> sections,
        VisioTextSectionSource? source, VisioTextStyle? style, bool character) {
        XElement? raw = sections.FirstOrDefault(section => character
            ? IsCharacterSection((string?)section.Attribute("N"))
            : IsParagraphSection((string?)section.Attribute("N")));
        if (raw != null) return raw;
        if (source == null) {
            if (style == null) return null;
            XElement modeled = CreateTextStyleSection(VisioNamespace, style, character);
            return modeled.Descendants(XName.Get("Cell", VisioNamespace)).Any() ? modeled : null;
        }
        XElement current = CreateTextStyleSection(source.Source.Name.NamespaceName, style, character);
        AlignTextSectionIdentity(current, source.Source);
        XElement output = new(source.Source);
        MergeModeledContentChanges(output, source.Baseline, current);
        PreserveUnmodeledCharacterStyleBits(output, source, current, character);
        ReplaceAssignedFont(output, current, style, character);
        // Inherited caches carry an empty row identity until a local edit adds
        // a cell. Keep that placeholder out of render-time ancestor selection.
        if (source.Inherited && !output.Descendants(XName.Get("Cell", output.Name.NamespaceName)).Any()) return null;
        return output;
    }

    private static VisioTextSectionSource CaptureTextSection(XElement source, VisioTextStyle? style, bool character, bool inherited = false) {
        XElement baseline = CreateTextStyleSection(source.Name.NamespaceName, style, character);
        AlignTextSectionIdentity(baseline, source);
        return new VisioTextSectionSource(source, baseline, inherited);
    }

    private static VisioTextSectionSource CaptureInheritedTextSection(VisioTextSectionSource? local,
        VisioTextSectionSource? master, VisioTextStyle? style, XNamespace ns, bool character) {
        if (local != null) return CaptureTextSection(local.Source, style, character, local.Inherited);
        XElement? masterRow = master?.Source.Elements(ns + "Row").SingleOrDefault();
        var row = masterRow == null ? new XElement(ns + "Row", new XAttribute("IX", "0")) :
            new XElement(ns + "Row", masterRow.Attributes().Where(a => a.Name == "N" || a.Name == "IX" || a.Name == "ID")
                .Select(a => new XAttribute(a)));
        var source = new XElement(ns + "Section", new XAttribute("N", character ? "Character" : "Paragraph"), row);
        return CaptureTextSection(source, style, character, inherited: true);
    }

    private static XElement CreateTextStyleSection(string ns, VisioTextStyle? style, bool character) {
        var document = new XDocument();
        using (XmlWriter writer = document.CreateWriter()) {
            if (character) WriteCharSection(writer, ns, style);
            else WriteParaSection(writer, ns, style);
        }
        // Clearing modeled formatting leaves the native row available to its text markers.
        return document.Root ?? new XElement(XName.Get("Section", ns),
            new XAttribute("N", character ? "Character" : "Paragraph"),
            new XElement(XName.Get("Row", ns), new XAttribute("IX", "0")));
    }

    private static void AlignTextSectionIdentity(XElement modeled, XElement source) {
        modeled.SetAttributeValue("N", (string?)source.Attribute("N"));
        XElement? nativeRow = source.Elements(source.Name.Namespace + "Row").SingleOrDefault();
        XElement? modelRow = modeled.Elements(modeled.Name.Namespace + "Row").SingleOrDefault();
        if (nativeRow == null || modelRow == null) return;
        foreach (string name in new[] { "N", "IX", "ID" })
            modelRow.SetAttributeValue(name, (string?)nativeRow.Attribute(name));
    }

    private static void WriteTextSectionSource(XmlWriter writer, string ns, VisioTextStyle? style,
        VisioTextSectionSource preserved, bool character) {
        XElement current = CreateTextStyleSection(ns, style, character);
        AlignTextSectionIdentity(current, preserved.Source);
        XElement output = new(preserved.Source);
        MergeModeledContentChanges(output, preserved.Baseline, current);
        PreserveUnmodeledCharacterStyleBits(output, preserved, current, character);
        ReplaceAssignedFont(output, current, style, character);
        // Inherited model values are rendering caches, not local overrides. Retain
        // the master's row identity only when an explicit edit adds a local cell.
        if (preserved.Inherited && !output.Descendants(XName.Get("Cell", ns)).Any()) return;
        using XmlReader reader = output.CreateReader();
        writer.WriteNode(reader, false);
    }

    private static void PreserveUnmodeledCharacterStyleBits(XElement output, VisioTextSectionSource source,
        XElement current, bool character) {
        if (!character) return;
        XNamespace ns = output.Name.Namespace;
        XElement? Cell(XElement section) => section.Elements(ns + "Row").Elements(ns + "Cell").FirstOrDefault(c => (string?)c.Attribute("N") == "Style");
        XElement? original = Cell(source.Source), baseline = Cell(source.Baseline), changed = Cell(current), target = Cell(output);
        if (!TryParseCellIntValue((string?)original?.Attribute("V"), out int native)) return;
        string? before = (string?)baseline?.Attribute("V"), after = (string?)changed?.Attribute("V");
        if (before == after) return;
        int modeled = 0;
        if (after != null && !TryParseCellIntValue(after, out modeled)) return;
        int combined = (native & ~15) | (modeled & 15);
        if (target == null && combined != 0) {
            // Clearing the supported flags must not clear producer-only flags.
            target = new XElement(original!);
            output.Elements(ns + "Row").Single().Add(target);
        }
        if (target == null) return;
        target.SetAttributeValue("V", combined);
        target.SetAttributeValue("F", null);
        target.SetAttributeValue("E", null);
        target.SetAttributeValue("Err", null);
    }

    // A public assignment replaces the native Font cell even when the family did not
    // change. Producer formulas/errors must not regain authority through a same-value edit.
    private static void ReplaceAssignedFont(XElement output, XElement current, VisioTextStyle? style, bool character) {
        if (!character || style?.FontFamilyAssigned != true) return;
        XNamespace ns = output.Name.Namespace;
        XElement? nativeRow = output.Elements(ns + "Row").SingleOrDefault();
        if (nativeRow == null) return;
        XElement? replacement = current.Elements(ns + "Row").Elements(ns + "Cell").FirstOrDefault(cell => (string?)cell.Attribute("N") == "Font");
        XElement? font = nativeRow.Elements(ns + "Cell").FirstOrDefault(cell => (string?)cell.Attribute("N") == "Font");
        if (font != null) { if (replacement == null) font.Remove(); else font.ReplaceWith(new XElement(replacement)); }
        else if (replacement != null) nativeRow.Add(new XElement(replacement));
    }
}
