using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;

namespace OfficeIMO.Visio;

public partial class VisioDocument {
    private void ImportStyleSheets(XElement? styleSheets, XNamespace ns, IReadOnlyDictionary<string, XElement?> nativeCells) {
        if (styleSheets == null) return;
        foreach (XAttribute attribute in styleSheets.Attributes().Where(ShouldPreserveStyleSheetsAttribute))
            AddMissingAttribute(PreservedStyleSheetsAttributes, attribute);
        foreach (XElement element in styleSheets.Elements().Where(ShouldPreserveStyleSheetsElement))
            if (!PreservedStyleSheetsElements.Any(existing => XNode.DeepEquals(existing, element)))
                PreservedStyleSheetsElements.Add(new XElement(element));

        foreach (XElement styleSheet in styleSheets.Elements(ns + "StyleSheet")) {
            string? id = (string?)styleSheet.Attribute("ID");
            if (string.IsNullOrWhiteSpace(id)) continue;
            id = NormalizeStyleSheetId(id!);
            if (!IsGeneratedStyleSheet(id)) {
                XElement? existing = PreservedAdditionalStyleSheets.FirstOrDefault(style =>
                    NormalizeStyleSheetId((string?)style.Attribute("ID") ?? string.Empty) == id);
                if (existing == null) {
                    PreservedAdditionalStyleSheets.Add(new XElement(styleSheet));
                    TransferImportedStyleCellMetadata(styleSheet, nativeCells, VisioNativeCellMetadata.Address(styleSheet));
                } else {
                    MergeImportedStyleElements(existing.Elements(), styleSheet.Elements(), existing.Add, nativeCells,
                        "StyleSheets/" + VisioNativeCellMetadata.Step(existing));
                }
                continue;
            }

            // A loaded/imported header owns its defaults, including absent attributes.
            // A synthesized generated style has no native header yet, so retain the first import's metadata.
            bool existingHeader = PreservedGeneratedStyleSheets.ContainsKey(id);
            PreservedStyleSheetData preserved = GetOrCreatePreservedStyleSheet(this, id);
            if (!existingHeader)
                foreach (XAttribute attribute in styleSheet.Attributes().Where(attribute => ShouldPreserveStyleSheetAttribute(attribute, id)))
                    AddMissingAttribute(preserved.Attributes, attribute);
            MergeImportedStyleElements(preserved.ChildElements,
                styleSheet.Elements().Where(element => ShouldPreserveStyleSheetElement(element, id)), preserved.ChildElements.Add, nativeCells,
                "StyleSheets/" + VisioNativeCellMetadata.Step(new XElement(ns + "StyleSheet", new XAttribute("ID", id))));
        }
    }

    // Native cell identities are finer than their containing section. Destination values
    // remain authoritative; only genuine additions carry the source's guarded metadata.
    private void MergeImportedStyleElements(IEnumerable<XElement> destination, IEnumerable<XElement> source,
        Action<XElement> append, IReadOnlyDictionary<string, XElement?> nativeCells, string destinationAddress) {
        XNamespace ns = VisioNamespace;
        foreach (XElement imported in source) {
            bool nativeIdentity = imported.Name == ns + "Section" || imported.Name == ns + "Row" || imported.Name == ns + "Cell";
            if (!nativeIdentity) {
                if (!destination.Any(existing => XNode.DeepEquals(existing, imported))) append(new XElement(imported));
                continue;
            }
            string identity = ImportedStyleElementIdentity(imported);
            XElement[] matches = destination.Where(existing => existing.Name == imported.Name &&
                ImportedStyleElementIdentity(existing) == identity).ToArray();
            if (matches.Length == 0) {
                append(new XElement(imported));
                TransferImportedStyleCellMetadata(imported, nativeCells, destinationAddress + "/" + VisioNativeCellMetadata.Step(imported));
            } else if (matches.Length == 1 && imported.Name != ns + "Cell") {
                XElement existing = matches[0];
                // Missing Del and other optional attributes are destination defaults, not vacant slots.
                MergeImportedStyleElements(existing.Elements(), imported.Elements(), existing.Add, nativeCells,
                    destinationAddress + "/" + VisioNativeCellMetadata.Step(existing));
            }
        }
    }

    // Imported named rows use their universal name, while the retained native IX is
    // still part of metadata addresses and the general row-identity edit boundary.
    private static string ImportedStyleElementIdentity(XElement element) =>
        element.Name == XName.Get("Row", VisioNamespace) && (string?)element.Attribute("N") is string name && name.Length > 0
            ? "Row[N=" + Uri.EscapeDataString(name) + "]" : VisioNativeCellMetadata.Step(element);
}
