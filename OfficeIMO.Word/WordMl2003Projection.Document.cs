using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word;

internal sealed partial class WordMl2003Projection {
    private static readonly XNamespace Office = "urn:schemas-microsoft-com:office:office";

    private void ValidateDocumentSingletons(XElement root) {
        foreach (XName name in new[] { Word + "styles", Word + "fonts", Word + "docPr", Office + "DocumentProperties" }) {
            if (root.Elements(name).Skip(1).Any())
                throw new InvalidDataException("WordML 2003 contains a duplicate " + name.LocalName + " declaration.");
        }
        if (root.Element(Word + "fonts")?.Elements(Word + "defaultFonts").Skip(1).Any() == true)
            throw new InvalidDataException("WordML 2003 contains a duplicate defaultFonts declaration.");
    }

    private void ReadDocumentParts(XElement root) {
        CheckAttributes(root, "macrosPresent embeddedObjPresent ocxPresent");
        IndexImages(root);
        foreach (XElement child in root.Elements()) {
            _token.ThrowIfCancellationRequested();
            if (child.Name == Word + "body") continue;
            if (child.Name == Word + "styles") ReadStyles(child, root.Element(Word + "fonts")?.Element(Word + "defaultFonts"));
            else if (child.Name == Word + "fonts") ReadFonts(child);
            else if (child.Name == Word + "docPr") ReadSettings(child);
            else if (child.Name == Office + "DocumentProperties") ReadMetadata(child);
            else if (child.Name == Word + "ignoreSubtree") {
                CheckAttributes(child, "val");
                // The producer explicitly excludes this namespace from Word's document semantics.
            } else Omit(child, "document-level content");
        }
        if (_main.StyleDefinitionsPart == null && root.Element(Word + "fonts")?.Element(Word + "defaultFonts") is { } defaults)
            ReadStyles(new XElement(Word + "styles"), defaults);
    }

    private void ReadStyles(XElement source, XElement? defaultFonts) {
        CheckAttributes(source, "");
        var styles = new Styles();
        if (defaultFonts != null) {
            XElement? mapped = MapProperties(defaultFonts, "rFonts");
            styles.DocDefaults = new DocDefaults(new RunPropertiesDefault(
                NormalizeMappedProperties(new RunPropertiesBaseStyle(new XElement(ModernWord + "rPr", mapped).ToString(SaveOptions.DisableFormatting)))));
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
            string? kind = (string?)child.Attribute(Word + "type");
            StyleValues type = kind switch {
                "paragraph" => StyleValues.Paragraph,
                "character" => StyleValues.Character,
                "table" => StyleValues.Table,
                "list" => StyleValues.Numbering,
                _ => throw new InvalidDataException("WordML 2003 contains an unsupported style type.")
            };
            var style = new Style { StyleId = id, Type = type };
            if (child.Attribute(Word + "default") is { } defaultAttribute) style.Default = IsTrue(defaultAttribute.Value);
            if (child.Attribute(Word + "customStyle") is { } customAttribute) style.CustomStyle = IsTrue(customAttribute.Value);
            foreach (XElement property in child.Elements()) {
                if (property.Name == Auxiliary + "uiName" || property.Name == Word + "rsid") continue;
                if (property.Name == Word + "name") style.StyleName = new StyleName { Val = ReadStyleReference(property) };
                else if (property.Name == Word + "basedOn") style.BasedOn = new BasedOn { Val = ReadStyleReference(property) };
                else if (property.Name == Word + "next") style.NextParagraphStyle = new NextParagraphStyle { Val = ReadStyleReference(property) };
                else if (property.Name == Word + "link") style.LinkedStyle = new LinkedStyle { Val = ReadStyleReference(property) };
                else if (property.Name == Word + "pPr") style.StyleParagraphProperties =
                    NormalizeMappedProperties(new StyleParagraphProperties(MapProperties(property, "pPr")!.ToString(SaveOptions.DisableFormatting)));
                else if (property.Name == Word + "rPr") style.StyleRunProperties =
                    NormalizeMappedProperties(new StyleRunProperties(MapProperties(property, "rPr")!.ToString(SaveOptions.DisableFormatting)));
                else if (property.Name == Word + "tblPr") style.StyleTableProperties =
                    NormalizeMappedProperties(new StyleTableProperties(MapProperties(property, "tblPr")!.ToString(SaveOptions.DisableFormatting)));
                else if (property.Name == Word + "trPr") style.TableStyleConditionalFormattingTableRowProperties =
                    NormalizeMappedProperties(new TableStyleConditionalFormattingTableRowProperties(MapProperties(property, "trPr")!.ToString(SaveOptions.DisableFormatting)));
                else if (property.Name == Word + "tcPr") style.StyleTableCellProperties =
                    NormalizeMappedProperties(new StyleTableCellProperties(MapProperties(property, "tcPr")!.ToString(SaveOptions.DisableFormatting)));
                else if (property.Name == Word + "tblStylePr") {
                    XElement? conditional = MapProperties(property, "tblStylePr");
                    if (conditional != null) style.AppendChild(NormalizeMappedProperties(new TableStyleProperties(conditional.ToString(SaveOptions.DisableFormatting))));
                }
                else Omit(property, "style properties");
            }
            styles.AppendChild(style);
        }
        ReserveRelationship(_main, createsPart: true);
        _main.AddNewPart<StyleDefinitionsPart>().Styles = styles;
    }

    private string ReadStyleReference(XElement source) {
        CheckAttributes(source, "val");
        if (source.HasElements) Omit(source, "nested style-reference content");
        return (string?)source.Attribute(Word + "val") ?? string.Empty;
    }

    private void ReadFonts(XElement source) {
        CheckAttributes(source, "");
        var fonts = new Fonts();
        foreach (XElement child in source.Elements()) {
            if (child.Name == Word + "defaultFonts") continue;
            if (child.Name != Word + "font") { Omit(child, "font-table content"); continue; }
            CheckAttributes(child, "name");
            var font = new XElement(ModernWord + "font", new XAttribute(ModernWord + "name", (string?)child.Attribute(Word + "name") ?? string.Empty));
            foreach (XElement property in child.Elements()) {
                XElement? mapped = MapProperties(property, property.Name.LocalName);
                if (mapped == null) continue;
                if (property.Name == Word + "panose-1") { mapped.Name = ModernWord + "panose1"; font.Add(mapped); }
                else if (property.Name == Word + "charset") font.Add(mapped);
                else if (property.Name == Word + "family") font.Add(mapped);
                else if (property.Name == Word + "pitch") font.Add(mapped);
                else if (property.Name == Word + "sig") Omit(property, "font signature metadata");
                else Omit(property, "font metadata");
            }
            fonts.AppendChild(NormalizeMappedProperties(new Font(font.ToString(SaveOptions.DisableFormatting))));
        }
        if (fonts.HasChildren) {
            ReserveRelationship(_main, createsPart: true);
            _main.AddNewPart<FontTablePart>().Fonts = fonts;
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
            var ordered = NormalizeMappedProperties(new Settings(settings.ToString(SaveOptions.DisableFormatting)));
            ReserveRelationship(_main, createsPart: true);
            _main.AddNewPart<DocumentSettingsPart>().Settings = ordered;
        }
    }

    private void ReadMetadata(XElement source) {
        CheckAttributes(source, "", Office);
        foreach (XElement child in source.Elements()) {
            CheckAttributes(child, "", Office);
            if (child.HasElements || child.Name.Namespace != Office) { Omit(child, "document properties"); continue; }
            string value = child.Value;
            if (ContainsName("Title Subject Author LastAuthor Keywords Description", child.Name.LocalName) && !_corePropertiesReserved) {
                ReserveRelationship(_main.OpenXmlPackage, createsPart: true);
                _corePropertiesReserved = true;
            }
            switch (child.Name.LocalName) {
                case "Title": _main.OpenXmlPackage.PackageProperties.Title = value; break;
                case "Subject": _main.OpenXmlPackage.PackageProperties.Subject = value; break;
                case "Author": _main.OpenXmlPackage.PackageProperties.Creator = value; break;
                case "LastAuthor": _main.OpenXmlPackage.PackageProperties.LastModifiedBy = value; break;
                case "Keywords": _main.OpenXmlPackage.PackageProperties.Keywords = value; break;
                case "Description": _main.OpenXmlPackage.PackageProperties.Description = value; break;
                default: Omit(child, "document properties"); break;
            }
        }
    }

    private void AppendHeaderFooter(SectionProperties properties, XElement source) {
        CheckAttributes(source, "type");
        HeaderFooterValues type = (string?)source.Attribute(Word + "type") switch {
            "odd" => HeaderFooterValues.Default, "even" => HeaderFooterValues.Even, "first" => HeaderFooterValues.First,
            _ => throw new InvalidDataException("WordML 2003 contains an invalid header or footer type.")
        };
        OpenXmlPart previous = _story;
        ReserveRelationship(_main, createsPart: true);
        try {
            if (source.Name == Word + "hdr") {
                HeaderPart part = _main.AddNewPart<HeaderPart>();
                _story = part;
                part.Header = new Header(ProjectBlocks(source));
                if (!part.Header.HasChildren) part.Header.AppendChild(new Paragraph());
                properties.PrependChild(new HeaderReference { Type = type, Id = _main.GetIdOfPart(part) });
            } else {
                FooterPart part = _main.AddNewPart<FooterPart>();
                _story = part;
                part.Footer = new Footer(ProjectBlocks(source));
                if (!part.Footer.HasChildren) part.Footer.AppendChild(new Paragraph());
                properties.PrependChild(new FooterReference { Type = type, Id = _main.GetIdOfPart(part) });
            }
        } finally { _story = previous; }
    }
}
