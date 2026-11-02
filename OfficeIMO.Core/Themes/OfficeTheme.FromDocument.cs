using System;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Xml;
using System.Xml.Linq;

namespace OfficeIMO;

public sealed partial class OfficeTheme {
    /// <summary>Reads the colour scheme and Latin heading/body fonts from a Word, Excel or PowerPoint package.</summary>
    /// <remarks>
    /// Accepts DOCX/DOTX, XLSX/XLTX and PPTX/POTX. PowerPoint uses the first declared slide master's theme.
    /// Spacing, type sizes and other composition defaults come from Modern; document content is not imported.
    /// External relationships are never followed. Encrypted packages must be decrypted before reading their theme.
    /// </remarks>
    /// <param name="path">Path to the document or template.</param>
    /// <returns>An immutable theme suitable for creating documents in any supported format.</returns>
    public static OfficeTheme FromDocument(string path) {
        if (string.IsNullOrWhiteSpace(path)) throw new ArgumentException("A document path is required.", nameof(path));
        string extension = Path.GetExtension(path).ToLowerInvariant();
        if (extension is not (".docx" or ".dotx" or ".xlsx" or ".xltx" or ".pptx" or ".potx"))
            throw new ArgumentException("Use a DOCX/DOTX, XLSX/XLTX or PPTX/POTX theme source.", nameof(path));
        using Stream input = File.OpenRead(path);
        using var package = new ZipArchive(input, ZipArchiveMode.Read);
        string mainPart = ReadThemeRelationship(package, string.Empty, "officeDocument");
        if (extension is ".pptx" or ".potx") {
            XDocument presentation = ReadThemePart(package, mainPart);
            XElement? master = presentation.Descendants().FirstOrDefault(element => element.Name.LocalName == "sldMasterId");
            string? masterId = master?.Attributes().FirstOrDefault(attribute => attribute.Name.LocalName == "id" &&
                attribute.Name.NamespaceName.Length > 0)?.Value;
            if (masterId == null) throw new InvalidDataException("The presentation has no declared slide master.");
            mainPart = ReadThemeRelationship(package, mainPart, "slideMaster", masterId);
        }
        XDocument nativeTheme = ReadThemePart(package, ReadThemeRelationship(package, mainPart, "theme"));
        XElement root = nativeTheme.Root ?? throw new InvalidDataException("The theme is empty.");
        XNamespace drawing = root.Name.Namespace;
        if (root.Name.LocalName != "theme" || (drawing.NamespaceName != "http://schemas.openxmlformats.org/drawingml/2006/main" &&
            drawing.NamespaceName != "http://purl.oclc.org/ooxml/drawingml/main"))
            throw new InvalidDataException("The package theme is not an Office DrawingML theme.");
        XElement elements = root.Element(drawing + "themeElements") ?? throw new InvalidDataException("Theme elements are missing.");
        XElement colors = elements.Element(drawing + "clrScheme") ?? throw new InvalidDataException("The theme colour scheme is missing.");
        OfficeColor Color(string slot) {
            XElement? color = colors.Element(drawing + slot)?.Elements().FirstOrDefault();
            string? hex = color?.Name == drawing + "srgbClr" ? (string?)color.Attribute("val") :
                color?.Name == drawing + "sysClr" ? (string?)color.Attribute("lastClr") : null;
            if (hex == null || hex.Length != 6 || hex.Any(character => !Uri.IsHexDigit(character)))
                throw new InvalidDataException("Theme colour " + slot + " has no literal RGB value.");
            return OfficeColor.ParseHex(hex);
        }
        var palette = new OfficeThemePalette(Color("lt1"), Color("dk1"), Color("dk2"),
            Color("accent1"), Color("accent2"), Color("accent3"), Color("accent4"), Color("accent5"), Color("accent6"),
            Modern.Palette.Success, Modern.Palette.Warning, Modern.Palette.Danger);
        XElement? fontScheme = elements.Element(drawing + "fontScheme");
        OfficeThemeFont Font(string slot, OfficeThemeFont fallback) {
            string? family = (string?)fontScheme?.Element(drawing + slot)?.Element(drawing + "latin")?.Attribute("typeface");
            return string.IsNullOrWhiteSpace(family) ? fallback : fallback.With(family: family!);
        }
        var fonts = new OfficeThemeFonts(Font("majorFont", Modern.Fonts.Heading), Font("minorFont", Modern.Fonts.Body));
        var surfaces = OfficeThemeSurfaces.FromPalette(palette, Modern.Surfaces.Surface, Modern.Surfaces.Rule);
        return Modern.WithName((string?)root.Attribute("name") ?? Path.GetFileNameWithoutExtension(path))
            .WithPalette(palette).WithFonts(fonts).WithSurfaces(surfaces)
            .WithTableStyle(Modern.TableStyle.With(headerFill: surfaces.Accent1Tint))
            .WithChartStyle(Modern.ChartStyle.With(axisFont: fonts.Body, titleFont: fonts.Heading));
    }

    private static string ReadThemeRelationship(ZipArchive package, string owner, string kind, string? id = null) {
        int slash = owner.LastIndexOf('/');
        string relationshipPart = owner.Length == 0 ? "_rels/.rels" :
            owner.Substring(0, slash + 1) + "_rels/" + owner.Substring(slash + 1) + ".rels";
        XDocument relationships = ReadThemePart(package, relationshipPart);
        XNamespace ns = "http://schemas.openxmlformats.org/package/2006/relationships";
        XElement relationship = relationships.Root?.Elements(ns + "Relationship").FirstOrDefault(element =>
            ((string?)element.Attribute("Type") == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/" + kind ||
             (string?)element.Attribute("Type") == "http://purl.oclc.org/ooxml/officeDocument/relationships/" + kind) &&
            (id == null || (string?)element.Attribute("Id") == id)) ??
            throw new InvalidDataException("The package has no " + kind + " relationship.");
        string? target = (string?)relationship.Attribute("Target");
        string? targetMode = (string?)relationship.Attribute("TargetMode");
        if ((targetMode != null && targetMode != "Internal") || string.IsNullOrWhiteSpace(target) ||
            target!.IndexOf('\\') >= 0 || target.IndexOf(':') >= 0 || target.IndexOf('#') >= 0 || target.IndexOf('?') >= 0)
            throw new InvalidDataException("The package theme relationship must target an internal part.");
        var root = new Uri("https://officeimo.invalid/");
        var resolved = new Uri(new Uri(root, owner), target);
        if (resolved.Scheme != root.Scheme || resolved.Host != root.Host)
            throw new InvalidDataException("The package theme relationship must target an internal part.");
        return Uri.UnescapeDataString(resolved.AbsolutePath.TrimStart('/'));
    }

    private static XDocument ReadThemePart(ZipArchive package, string part) {
        ZipArchiveEntry entry = package.GetEntry(part) ?? throw new InvalidDataException("The package part is missing: " + part);
        const long maxCharacters = 2 * 1024 * 1024;
        if (entry.Length > maxCharacters) throw new InvalidDataException("The theme source part exceeds the 2 MiB limit: " + part);
        using Stream input = entry.Open();
        using XmlReader reader = XmlReader.Create(input, new XmlReaderSettings {
            DtdProcessing = DtdProcessing.Prohibit, XmlResolver = null, MaxCharactersInDocument = maxCharacters
        });
        return XDocument.Load(reader);
    }
}
