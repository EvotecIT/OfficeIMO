using System.Xml.Linq;
using System.Globalization;
using OfficeIMO.Core.Internal.OpenXml;

namespace OfficeIMO.Word;

internal sealed partial class WordMl2003Projection {
    private static readonly XNamespace Office = "urn:schemas-microsoft-com:office:office";

    private void ValidateDocumentSingletons(XElement root) {
        foreach (XName name in new[] { Word + "styles", Word + "fonts", Word + "lists", Word + "docPr", Office + "DocumentProperties" }) {
            if (root.Elements(name).Skip(1).Any())
                throw new InvalidDataException("WordML 2003 contains a duplicate " + name.LocalName + " declaration.");
        }
        if (root.Element(Word + "fonts")?.Elements(Word + "defaultFonts").Skip(1).Any() == true)
            throw new InvalidDataException("WordML 2003 contains a duplicate defaultFonts declaration.");
    }

    private void ReadDocumentParts(XElement root) {
        CheckAttributes(root, "macrosPresent embeddedObjPresent ocxPresent");
        IndexImages(root);
        if (root.Element(Word + "lists") is { } lists) ReadLists(lists);
        foreach (XElement child in root.Elements()) {
            _token.ThrowIfCancellationRequested();
            if (child.Name == Word + "body" || child.Name == Word + "lists") continue;
            if (child.Name == Word + "styles") ReadStyles(child, root.Element(Word + "fonts")?.Element(Word + "defaultFonts"));
            else if (child.Name == Word + "fonts") ReadFonts(child);
            else if (child.Name == Word + "docPr") ReadSettings(child);
            else if (child.Name == Office + "DocumentProperties") ReadMetadata(child);
            else if (child.Name == Word + "ignoreSubtree") {
                CheckAttributes(child, "val");
                // The producer explicitly excludes this namespace from Word's document semantics.
            } else Omit(child, "document-level content");
        }
        if (!_package.PartExists("/word/styles.xml") && root.Element(Word + "fonts")?.Element(Word + "defaultFonts") is { } defaults)
            ReadStyles(new XElement(Word + "styles"), defaults);
    }

    private void ReadStyles(XElement source, XElement? defaultFonts) {
        CheckAttributes(source, "");
        OfficeXmlTree styles = AddMainPart("styles", "styles", out _);
        if (defaultFonts != null) {
            int defaults = AppendWord(styles, styles.Root, "docDefaults");
            int runDefault = AppendWord(styles, defaults, "rPrDefault");
            var mapped = new XElement(ModernWord + "rPr", MapProperties(defaultFonts, "rFonts"));
            NormalizeMappedProperties(mapped, "defaultRun");
            AppendMapped(styles, runDefault, mapped);
        }
        var ids = new HashSet<string>(StringComparer.Ordinal);
        foreach (XElement child in source.Elements()) {
            if (child.Name == Word + "versionOfBuiltInStylenames" || child.Name == Word + "latentStyles") {
                Omit(child, "latent or built-in style metadata");
                continue;
            }
            if (child.Name != Word + "style") { Omit(child, "style-library content"); continue; }
            CheckAttributes(child, "type styleId default customStyle");
            string? id = (string?)child.Attribute(Word + "styleId");
            if (id == null || id.Length == 0 || !ids.Add(id))
                throw new InvalidDataException("WordML 2003 contains a missing or duplicate style identifier.");
            string type = (string?)child.Attribute(Word + "type") switch {
                "paragraph" => "paragraph",
                "character" => "character",
                "table" => "table",
                "list" => "numbering",
                _ => throw new InvalidDataException("WordML 2003 contains an unsupported style type.")
            };
            int style = AppendWord(styles, styles.Root, "style");
            styles.SetAttribute(style, W, "type", type);
            styles.SetAttribute(style, W, "styleId", id);
            if (child.Attribute(Word + "default") is { } defaultAttribute) styles.SetAttribute(style, W, "default", BooleanToken(IsTrue(defaultAttribute.Value)));
            if (child.Attribute(Word + "customStyle") is { } customAttribute) styles.SetAttribute(style, W, "customStyle", BooleanToken(IsTrue(customAttribute.Value)));
            foreach (XElement property in child.Elements()) {
                if (property.Name == Auxiliary + "uiName" || property.Name == Word + "rsid") continue;
                string name = property.Name.LocalName;
                if (property.Name.Namespace == Word && ContainsName("name basedOn next link", name)) {
                    int reference = styles.CreateElement(W, name);
                    styles.SetAttribute(reference, W, "val", ReadStyleReference(property));
                    ReplaceWordChild(styles, style, reference, StyleOrder);
                } else if (property.Name.Namespace == Word && ContainsName("pPr rPr tblPr trPr tcPr", name)) {
                    ReplaceWordChild(styles, style, ProjectProperties(styles, property, name, context: "style"), StyleOrder);
                } else if (property.Name == Word + "tblStylePr") {
                    XElement? conditional = MapProperties(property, "tblStylePr");
                    if (conditional != null) {
                        NormalizeMappedProperties(conditional);
                        AppendMapped(styles, style, conditional);
                    }
                } else Omit(property, "style properties");
            }
        }
    }

    private static readonly OfficeXmlChildOrder StyleOrder = WordChildOrder("name basedOn next link pPr rPr tblPr trPr tcPr tblStylePr");

    private static string BooleanToken(bool value) => value ? "true" : "false";

    private string ReadStyleReference(XElement source) {
        CheckAttributes(source, "val");
        if (source.HasElements) Omit(source, "nested style-reference content");
        return (string?)source.Attribute(Word + "val") ?? string.Empty;
    }

    private void ReadFonts(XElement source) {
        CheckAttributes(source, "");
        OfficeXmlTree? fonts = null;
        foreach (XElement child in source.Elements()) {
            if (child.Name == Word + "defaultFonts") continue;
            if (child.Name != Word + "font") { Omit(child, "font-table content"); continue; }
            CheckAttributes(child, "name");
            var font = new XElement(ModernWord + "font", new XAttribute(ModernWord + "name", (string?)child.Attribute(Word + "name") ?? string.Empty));
            foreach (XElement property in child.Elements()) {
                XElement? mapped = MapProperties(property, property.Name.LocalName);
                if (mapped == null) continue;
                if (property.Name.Namespace == Word && ContainsName("panose-1 charset family pitch", property.Name.LocalName)) font.Add(mapped);
                else if (property.Name == Word + "sig") Omit(property, "font signature metadata");
                else Omit(property, "font metadata");
            }
            NormalizeMappedProperties(font);
            fonts ??= AddMainPart("fontTable", "fonts", out _);
            AppendMapped(fonts, fonts.Root, font);
        }
    }

    private void ReadSettings(XElement source) {
        CheckAttributes(source, "");
        var settings = new XElement(ModernWord + "settings");
        HashSet<string> supported = Names("view zoom doNotEmbedSystemFonts defaultTabStop punctuationKerning characterSpacingControl evenAndOddHeaders mirrorMargins displayBackgroundShape doNotHyphenateCaps autoHyphenation consecutiveHyphenLimit hyphenationZone");
        foreach (XElement child in source.Elements()) {
            if (child.Name.Namespace != Word || !supported.Contains(child.Name.LocalName)) {
                Omit(child, "document settings");
                continue;
            }
            XElement? mapped = MapProperties(child, child.Name.LocalName);
            if (mapped == null) continue;
            if (child.Name == Word + "doNotEmbedSystemFonts" || child.Name == Word + "punctuationKerning") {
                mapped.Name = ModernWord + (child.Name.LocalName == "doNotEmbedSystemFonts" ? "embedSystemFonts" : "noPunctuationKerning");
                mapped.SetAttributeValue(ModernWord + "val", !IsTrue((string?)child.Attribute(Word + "val") ?? "on"));
            }
            settings.Add(mapped);
        }
        if (settings.HasElements) {
            NormalizeMappedProperties(settings);
            OfficeXmlTree part = AddMainPart("settings", "settings", out _);
            foreach (XElement setting in settings.Elements()) AppendMapped(part, part.Root, setting);
        }
    }

    private void ReadMetadata(XElement source) {
        CheckAttributes(source, "", Office);
        var values = new Dictionary<string, string>(StringComparer.Ordinal);
        foreach (XElement child in source.Elements()) {
            CheckAttributes(child, "", Office);
            if (child.HasElements || child.Name.Namespace != Office ||
                !ContainsName("Title Subject Author LastAuthor Keywords Description", child.Name.LocalName)) {
                Omit(child, "document properties");
                continue;
            }
            if (values.Count == 0) ReserveRelationship(null, createsPart: true);
            values[child.Name.LocalName] = child.Value;
        }
        if (values.Count == 0) return;
        _package.AddRelationship(null, CorePropertiesRelationshipType, CorePropertiesPart);
        OfficeXmlTree core = _package.CreateXmlPart(CorePropertiesPart, CorePropertiesContentType, OfficeXmlNamespaces.CoreProperties, "coreProperties");
        Add("Title", "title", OfficeXmlNamespaces.DublinCore);
        Add("Subject", "subject", OfficeXmlNamespaces.DublinCore);
        Add("Keywords", "keywords", OfficeXmlNamespaces.CoreProperties);
        Add("Description", "description", OfficeXmlNamespaces.DublinCore);
        Add("Author", "creator", OfficeXmlNamespaces.DublinCore);
        Add("LastAuthor", "lastModifiedBy", OfficeXmlNamespaces.CoreProperties);

        void Add(string sourceName, string name, int ns) {
            if (!values.TryGetValue(sourceName, out string? text)) return;
            int node = core.CreateElement(ns, name);
            core.SetInnerText(node, text);
            core.Append(core.Root, node);
        }
    }

    private void AppendHeaderFooter(int properties, XElement source) {
        CheckAttributes(source, "type");
        string type = (string?)source.Attribute(Word + "type") switch {
            "odd" => "default", "even" => "even", "first" => "first",
            _ => throw new InvalidDataException("WordML 2003 contains an invalid header or footer type.")
        };
        Story previous = _story;
        bool header = source.Name == Word + "hdr";
        string kind = header ? "header" : "footer";
        string name = kind + (header ? ++_headerCount : ++_footerCount).ToString(CultureInfo.InvariantCulture);
        OfficeXmlTree tree = AddMainPart(name, header ? "hdr" : "ftr", out string id, kind);
        try {
            _story = new Story("/word/" + name + ".xml", tree);
            foreach (int block in ProjectBlocks(source)) tree.Append(tree.Root, block);
            if (tree.FirstChildOf(tree.Root) == NoNode) AppendWord(tree, tree.Root, "p");
        } finally { _story = previous; }
        OfficeXmlTree document = previous.Tree;
        int reference = document.CreateElement(W, kind + "Reference");
        document.SetAttribute(reference, W, "type", type);
        document.SetAttribute(reference, R, "id", id);
        document.InsertBefore(properties, reference, document.FirstChildOf(properties));
    }
}
