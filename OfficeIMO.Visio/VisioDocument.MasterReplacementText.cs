using System;
using System.Linq;
using System.Xml.Linq;

namespace OfficeIMO.Visio;

public partial class VisioDocument {
    /// <summary>Prepares local text rows so a changed master cannot alter retained cached formatting.</summary>
    internal static Action PrepareRetainedMasterText(VisioShape shape, VisioNativeTextStyleResolver resolver,
        IReadOnlyDictionary<VisioShape, IReadOnlyDictionary<string, string>> inheritedReferences) {
        XNamespace ns = VisioNamespace;
        bool changedText = shape.Text != shape.PreservedTextValue;
        XElement Prepare(bool character) {
            VisioTextSectionSource? source = character ? shape.CharacterSectionSource : shape.ParagraphSectionSource;
            XElement? local = GetRenderTextSection(shape.PreservedNonGeometrySections, source, shape.TextStyle, character);
            XElement? effective = resolver.ResolveNative(shape, character, changedText);
            XElement output = local == null ? new XElement(ns + "Section", new XAttribute("N", character ? "Character" : "Paragraph")) : new XElement(local);
            if (VisioNativeTextStyleResolver.HasDeletion(output)) return output;
            foreach (XElement inherited in effective?.Elements(ns + "Row") ?? Enumerable.Empty<XElement>()) {
                string index = VisioNativeTextStyleResolver.RowIndex((string?)inherited.Attribute("IX"));
                XElement? row = output.Elements(ns + "Row").FirstOrDefault(r => VisioNativeTextStyleResolver.RowIndex((string?)r.Attribute("IX")) == index);
                if (row == null) { row = new XElement(ns + "Row", inherited.Attributes().Select(a => new XAttribute(a))); output.Add(row); }
                foreach (XElement cell in inherited.Elements(ns + "Cell")) {
                    string? name = (string?)cell.Attribute("N");
                    if (!row.Elements(ns + "Cell").Any(c => (string?)c.Attribute("N") == name)) {
                        var copy = new XElement(cell);
                        if (cell.Annotation<VisioShape>() is VisioShape owner && inheritedReferences.TryGetValue(owner, out var references)
                            && copy.Attribute("F") is XAttribute formula)
                            formula.Value = VisioShapeFormulaReferences.Rewrite(formula.Value, references)!;
                        row.Add(copy);
                    }
                }
            }
            foreach (XElement cell in output.Descendants(ns + "Cell")) {
                if (!string.Equals((string?)cell.Attribute("F"), "Inh", StringComparison.OrdinalIgnoreCase)) continue;
                if (cell.Attribute("V") == null) throw new NotSupportedException("Retained inherited text cells require a cached value when replacing the master.");
                cell.Attribute("F")!.Remove();
            }
            if (!output.Descendants(ns + "Cell").Any()) output.SetAttributeValue("Del", "1");
            return output;
        }
        XElement character = Prepare(true), paragraph = Prepare(false);
        VisioTextSectionSource? charSource = character.Elements(ns + "Row").Count() <= 1 ? CaptureTextSection(character, shape.TextStyle, character: true) : null;
        VisioTextSectionSource? paraSource = paragraph.Elements(ns + "Row").Count() <= 1 ? CaptureTextSection(paragraph, shape.TextStyle, character: false) : null;
        return () => {
            // Complex rows remain native-owned; single rows retain the shared typed edit merge.
            foreach (XElement section in shape.PreservedNonGeometrySections.Where(section =>
                IsCharacterSection((string?)section.Attribute("N")) || IsParagraphSection((string?)section.Attribute("N"))).ToArray())
                shape.PreservedNonGeometrySections.Remove(section);
            if (character.Elements(ns + "Row").Count() > 1) { shape.PreservedNonGeometrySections.Add(character); shape.CharacterSectionSource = null; }
            else shape.CharacterSectionSource = charSource;
            if (paragraph.Elements(ns + "Row").Count() > 1) { shape.PreservedNonGeometrySections.Add(paragraph); shape.ParagraphSectionSource = null; }
            else shape.ParagraphSectionSource = paraSource;
        };
    }
}
