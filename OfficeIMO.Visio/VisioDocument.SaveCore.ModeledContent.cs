using System;
using System.Linq;
using System.Xml.Linq;

namespace OfficeIMO.Visio;

public partial class VisioDocument {
    // Apply only observed model edits; untouched native formulas and metadata remain authoritative.
    private static void MergeModeledContentChanges(XElement source, XElement baseline, XElement current) {
        if (XNode.DeepEquals(baseline, current)) return;
        PreserveMissingMasterLocalPins(source, baseline, current);
        foreach (XName name in baseline.Attributes().Concat(current.Attributes()).Where(a => !a.IsNamespaceDeclaration).Select(a => a.Name).Distinct()) {
            string? before = (string?)baseline.Attribute(name), after = (string?)current.Attribute(name);
            if (before != after) source.SetAttributeValue(name, after);
        }

        foreach (XElement previous in baseline.Elements()) {
            XElement? updated = MatchingModeledContentElement(current, previous);
            XElement? original = MatchingModeledContentElement(source, previous);
            if (updated == null) { original?.Remove(); continue; }
            if (XNode.DeepEquals(previous, updated)) continue;
            if (previous.Name.LocalName == "Shapes" && original != null) {
                // Model order and membership are authoritative for the editable child collection.
                var children = updated.Elements().Select(child => {
                    XElement? oldChild = MatchingModeledContentElement(previous, child);
                    XElement? sourceChild = MatchingModeledContentElement(original, child);
                    if (oldChild == null || sourceChild == null) return new XElement(child);
                    var merged = new XElement(sourceChild);
                    MergeModeledContentChanges(merged, oldChild, child);
                    return merged;
                }).ToArray();
                original.ReplaceNodes(children);
            } else if (original != null && (previous.Name.LocalName == "Section" || previous.Name.LocalName == "Row")) {
                MergeModeledContentChanges(original, previous, updated);
            } else if (original != null) {
                // Replace a changed cell/text value through the canonical writer. Keeping
                // the source formula on an explicitly changed modeled cell would undo that edit.
                original.ReplaceWith(new XElement(updated));
            } else {
                source.Add(new XElement(updated));
            }
        }
        foreach (XElement added in current.Elements().Where(e => MatchingModeledContentElement(baseline, e) == null))
            source.Add(new XElement(added));
    }

    private static void PreserveMissingMasterLocalPins(XElement source, XElement baseline, XElement current) {
        if (source.Name.LocalName != "Shape") return;
        XNamespace ns = source.Name.Namespace;
        XElement? Cell(XElement shape, string name) => shape.Elements(ns + "Cell").FirstOrDefault(cell => (string?)cell.Attribute("N") == name);
        foreach (var pair in new[] { (Extent: "Width", Pin: "LocPinX"), (Extent: "Height", Pin: "LocPinY") }) {
            // A missing native pin defaults to half the extent when reloaded. After an extent
            // edit the model still holds its original pin, so make that value explicit rather
            // than silently moving the saved child tree. Existing pin formulas remain untouched.
            if (Cell(source, pair.Pin) == null && !XNode.DeepEquals(Cell(baseline, pair.Extent), Cell(current, pair.Extent)) &&
                Cell(current, pair.Pin) is XElement pin) source.Add(new XElement(pin));
        }
    }

    private static XElement? MatchingModeledContentElement(XElement parent, XElement sample) => parent.Elements(sample.Name).FirstOrDefault(element =>
        (string?)element.Attribute("N") == (string?)sample.Attribute("N") &&
        (string?)element.Attribute("IX") == (string?)sample.Attribute("IX") &&
        (string?)element.Attribute("ID") == (string?)sample.Attribute("ID"));
}
