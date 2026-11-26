using OfficeIMO.Core.Internal.OpenXml;

namespace OfficeIMO.Word.OpenXml;

/// <summary>
/// The changes a load and save through the SDK path make to a document the owned layer holds, applied to the owned
/// package so both paths write the same XML: a body without final section properties gets Word's default ones, section
/// breaks without a revision identifier get one, the final section properties move to the end of the body, and
/// absolute relationship targets become relative, as <see cref="WordPackageCompatibility"/> writes them.
/// </summary>
internal static class WordOwnedSave {
    private const int W = OfficeXmlNamespaces.WordprocessingML;
    private const string PackageRelationshipsPart = "/_rels/.rels";
    private const string MainDocumentRelationshipsPart = "/word/_rels/document.xml.rels";

    /// <summary>
    /// Why the save has to go through the SDK, or null when the owned layer writes what the SDK path would. Tables are
    /// repaired and their grids normalised on that path's save, which the owned layer doesn't do yet.
    /// </summary>
    internal static string? FindSaveBlocker(WordPackageModel model) {
        if (ContainsElement(model.Document, model.Document.Root, "tbl")) return "tables in the document body";
        // The SDK path's save rewrites custom properties, document variables and bibliography sources from what it
        // loaded, which changes their XML; until those move, such documents save through it.
        if (HasPackageRelationship(model, CustomPropertiesRelationshipType)) return "custom document properties";
        if (model.Package.PartExists("/[Content_Types].xml") && DeclaresBibliography(model)) return "bibliography sources";
        foreach (string part in RelatedParts(model, SettingsRelationshipType)) {
            OfficeXmlTree settings = model.Package.GetXml(part);
            if (settings.FindChild(settings.Root, W, "docVars") != OfficeXmlTree.NoNode) return "document variables";
        }
        // The same parts the SDK path repairs tables in.
        foreach (string part in RelatedParts(model, HeaderRelationshipType, FooterRelationshipType, FootnotesRelationshipType, EndnotesRelationshipType)) {
            OfficeXmlTree tree = model.Package.GetXml(part);
            if (ContainsElement(tree, tree.Root, "tbl")) return "tables in a header, footer, footnote or endnote";
        }
        return null;
    }

    private const string HeaderRelationshipType = OfficeXmlNamespaces.TransitionalRelationshipTypePrefix + "header";
    private const string FooterRelationshipType = OfficeXmlNamespaces.TransitionalRelationshipTypePrefix + "footer";
    private const string FootnotesRelationshipType = OfficeXmlNamespaces.TransitionalRelationshipTypePrefix + "footnotes";
    private const string EndnotesRelationshipType = OfficeXmlNamespaces.TransitionalRelationshipTypePrefix + "endnotes";
    private const string SettingsRelationshipType = OfficeXmlNamespaces.TransitionalRelationshipTypePrefix + "settings";
    private const string CustomPropertiesRelationshipType = OfficeXmlNamespaces.TransitionalRelationshipTypePrefix + "custom-properties";
    private const string BibliographyContentType = "application/bibliography+xml";

    private static bool HasPackageRelationship(WordPackageModel model, string type) {
        if (!model.Package.PartExists("/_rels/.rels")) return false;
        OfficeXmlTree relationships = model.Package.GetRelationships(null);
        for (int child = relationships.FirstChildOf(relationships.Root); child != OfficeXmlTree.NoNode; child = relationships.NextSiblingOf(child)) {
            if (relationships.KindOf(child) != OfficeXmlTreeNodeKind.Element) continue;
            string? value = relationships.GetAttribute(child, OfficeXmlNamespaces.None, "Type");
            if (value != null && OfficeXmlNamespaces.NormalizeRelationshipType(value) == type) return true;
        }
        return false;
    }

    private static bool DeclaresBibliography(WordPackageModel model) {
        OfficeXmlTree types = model.Package.ContentTypes;
        for (int child = types.FirstChildOf(types.Root); child != OfficeXmlTree.NoNode; child = types.NextSiblingOf(child)) {
            if (types.KindOf(child) == OfficeXmlTreeNodeKind.Element
                && types.AttributeEquals(child, OfficeXmlNamespaces.None, "ContentType", BibliographyContentType)) return true;
        }
        return false;
    }

    private const string NumberingRelationshipType = OfficeXmlNamespaces.TransitionalRelationshipTypePrefix + "numbering";

    private static bool ContainsElement(OfficeXmlTree tree, int root, string localName) {
        for (int child = tree.FirstChildOf(root); child != OfficeXmlTree.NoNode; child = tree.NextSiblingOf(child)) {
            if (tree.KindOf(child) != OfficeXmlTreeNodeKind.Element) continue;
            if (tree.IsElement(child, W, localName) || ContainsElement(tree, child, localName)) return true;
        }
        return false;
    }

    private static IEnumerable<string> RelatedParts(WordPackageModel model, params string[] types) {
        // Asking the package for a relationships part it lacks would add an empty one.
        if (!model.Package.PartExists(OfficeOpcPartNames.GetRelationshipsPartName(model.MainPartName))) yield break;
        OfficeXmlTree relationships = model.Package.GetRelationships(model.MainPartName);
        for (int child = relationships.FirstChildOf(relationships.Root); child != OfficeXmlTree.NoNode; child = relationships.NextSiblingOf(child)) {
            if (relationships.KindOf(child) != OfficeXmlTreeNodeKind.Element) continue;
            string? type = relationships.GetAttribute(child, OfficeXmlNamespaces.None, "Type");
            if (type == null || Array.IndexOf(types, OfficeXmlNamespaces.NormalizeRelationshipType(type)) < 0) continue;
            if (string.Equals(relationships.GetAttribute(child, OfficeXmlNamespaces.None, "TargetMode"), "External", StringComparison.Ordinal)) continue;
            string? id = relationships.GetAttribute(child, OfficeXmlNamespaces.None, "Id");
            string? target = id == null ? null : model.Package.GetRelationshipTarget(model.MainPartName, id);
            if (target != null && model.Package.PartExists(target)) yield return target;
        }
    }

    internal static void PrepareForSave(WordPackageModel model) {
        OrderNumbering(model);
        EnsureBodySectionProperties(model);
        EnsureSectionRevisionIds(model);
        MoveFinalSectionPropertiesToEnd(model);
        NormalizeRelationshipTargets(model, PackageRelationshipsPart, prefix: "/", strip: 1);
        NormalizeRelationshipTargets(model, MainDocumentRelationshipsPart, prefix: "/word/", strip: 6);
    }

    /// <summary>
    /// Appends Word's default section properties to a body that has none, as the SDK path does when it loads such a
    /// document: Letter paper with Normal margins and a new revision identifier.
    /// </summary>
    internal static void EnsureBodySectionProperties(WordPackageModel model) {
        OfficeXmlTree document = model.Document;
        int body = model.Body;
        if (document.FindChild(body, W, "sectPr") != OfficeXmlTree.NoNode) return;
        int sectPr = document.CreateElement(W, "sectPr");
        document.SetAttribute(sectPr, W, "rsidR", WordHeadersAndFooters.GenerateRsid());
        int pageSize = document.CreateElement(W, "pgSz");
        document.SetAttribute(pageSize, W, "w", "12240");
        document.SetAttribute(pageSize, W, "h", "15840");
        document.SetAttribute(pageSize, W, "code", "1");
        document.Append(sectPr, pageSize);
        int margins = document.CreateElement(W, "pgMar");
        document.SetAttribute(margins, W, "top", "1440");
        document.SetAttribute(margins, W, "right", "1440");
        document.SetAttribute(margins, W, "bottom", "1440");
        document.SetAttribute(margins, W, "left", "1440");
        document.SetAttribute(margins, W, "header", "720");
        document.SetAttribute(margins, W, "footer", "720");
        document.SetAttribute(margins, W, "gutter", "0");
        document.Append(sectPr, margins);
        document.Append(body, sectPr);
        model.MarkChanged();
    }

    // The SDK path gives each section break it loads a revision identifier when it has none.
    private static void EnsureSectionRevisionIds(WordPackageModel model) {
        OfficeXmlTree document = model.Document;
        for (int child = document.FirstChildOf(model.Body); child != OfficeXmlTree.NoNode; child = document.NextSiblingOf(child)) {
            if (!document.IsElement(child, W, "p")) continue;
            int properties = document.FindChild(child, W, "pPr");
            if (properties == OfficeXmlTree.NoNode) continue;
            int sectPr = document.FindChild(properties, W, "sectPr");
            if (sectPr == OfficeXmlTree.NoNode || document.TryGetAttribute(sectPr, W, "rsidR", out _)) continue;
            document.SetAttribute(sectPr, W, "rsidR", WordHeadersAndFooters.GenerateRsid());
            model.MarkChanged();
        }
    }

    private static void MoveFinalSectionPropertiesToEnd(WordPackageModel model) {
        OfficeXmlTree document = model.Document;
        int body = model.Body;
        int last = LastElementChild(document, body);
        if (last != OfficeXmlTree.NoNode && document.IsElement(last, W, "sectPr")) return;
        int sectPr = OfficeXmlTree.NoNode;
        for (int child = document.FirstChildOf(body); child != OfficeXmlTree.NoNode; child = document.NextSiblingOf(child)) {
            if (document.IsElement(child, W, "sectPr")) sectPr = child;
        }
        if (sectPr == OfficeXmlTree.NoNode) return;
        document.Remove(sectPr);
        document.Append(body, sectPr);
        model.MarkChanged();
    }

    /// <summary>
    /// Puts picture bullets, abstract numberings and numbering instances in schema order when they are out of it, as
    /// the SDK path's save does. Other elements, such as <c>w:numIdMacAtCleanup</c>, stay after them; the SDK path drops
    /// those, Word keeps them.
    /// </summary>
    private static void OrderNumbering(WordPackageModel model) {
        string? part = RelatedParts(model, NumberingRelationshipType).FirstOrDefault();
        if (part == null) return;
        OfficeXmlTree numbering = model.Package.GetXml(part);
        int root = numbering.Root;
        int previous = 0;
        bool ordered = true;
        for (int child = numbering.FirstChildOf(root); child != OfficeXmlTree.NoNode; child = numbering.NextSiblingOf(child)) {
            if (numbering.KindOf(child) != OfficeXmlTreeNodeKind.Element) continue;
            int group = NumberingGroup(numbering, child);
            if (group < 0) group = 3;
            if (group < previous) {
                ordered = false;
                break;
            }
            previous = group;
        }
        if (ordered) return;
        var groups = new List<int>[] { new(), new(), new(), new() };
        for (int child = numbering.FirstChildOf(root); child != OfficeXmlTree.NoNode; child = numbering.NextSiblingOf(child)) {
            if (numbering.KindOf(child) != OfficeXmlTreeNodeKind.Element) continue;
            int group = NumberingGroup(numbering, child);
            groups[group < 0 ? 3 : group].Add(child);
        }
        foreach (List<int> group in groups) {
            foreach (int child in group) {
                numbering.Remove(child);
                numbering.Append(root, child);
            }
        }
        model.MarkChanged();
    }

    private static int NumberingGroup(OfficeXmlTree tree, int node) {
        if (tree.KindOf(node) != OfficeXmlTreeNodeKind.Element || tree.NamespaceOf(node) != W) return -1;
        return tree.LocalNameOf(node) switch {
            "numPicBullet" => 0,
            "abstractNum" => 1,
            "num" => 2,
            _ => -1
        };
    }

    internal static int LastElementChild(OfficeXmlTree tree, int parent) {
        for (int child = tree.LastChildOf(parent); child != OfficeXmlTree.NoNode; child = tree.PreviousSiblingOf(child)) {
            if (tree.KindOf(child) == OfficeXmlTreeNodeKind.Element) return child;
        }
        return OfficeXmlTree.NoNode;
    }

    private static void NormalizeRelationshipTargets(WordPackageModel model, string relationshipsPart, string prefix, int strip) {
        if (!model.Package.PartExists(relationshipsPart)) return;
        OfficeXmlTree relationships = model.Package.GetXml(relationshipsPart);
        int root = relationships.Root;
        for (int child = relationships.FirstChildOf(root); child != OfficeXmlTree.NoNode; child = relationships.NextSiblingOf(child)) {
            if (relationships.KindOf(child) != OfficeXmlTreeNodeKind.Element) continue;
            // An external target is a location outside the package, not a part name.
            if (relationships.AttributeEquals(child, OfficeXmlNamespaces.None, "TargetMode", "External")) continue;
            if (!relationships.TryGetAttribute(child, OfficeXmlNamespaces.None, "Target", out string target)
                || !target.StartsWith(prefix, StringComparison.Ordinal)) continue;
            relationships.SetAttribute(child, OfficeXmlNamespaces.None, "Target", prefix.Length == 1 ? target.TrimStart('/') : target.Substring(strip));
            model.MarkChanged();
        }
    }
}
