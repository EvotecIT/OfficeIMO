using System.Threading;
using System.Globalization;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word;

/// <summary>Projects the supported Word 2003 XML vocabulary into ordinary editable Word parts.</summary>
internal sealed class WordMl2003Projection {
    private static readonly XNamespace Word = WordXmlInterchange.WordMl;
    private static readonly XNamespace Auxiliary = "http://schemas.microsoft.com/office/word/2003/auxHint";
    private readonly CancellationToken _token;
    private readonly Dictionary<XName, OfficeDiagnostic> _omissions = new();

    private WordMl2003Projection(CancellationToken token) {
        _token = token;
    }

    internal static IReadOnlyList<OfficeDiagnostic> WritePackage(XElement root, Stream destination,
        OfficePackageSecurityOptions policy, CancellationToken token) {
        var projection = new WordMl2003Projection(token);
        XElement? body = root.Element(Word + "body");
        if (body == null || body.ElementsAfterSelf(Word + "body").Any()) {
            throw new InvalidDataException("WordML 2003 must contain exactly one Word body.");
        }
        projection.AdmitSource(root, policy);
        foreach (XElement child in root.Elements().Where(element => element != body)) {
            projection.Omit(child, "document-level settings or content");
        }

        using var document = WordprocessingDocument.Create(destination, WordprocessingDocumentType.Document);
        var output = new Body();
        SectionProperties? pendingSection = null;
        foreach (OpenXmlElement block in projection.ProjectBlocks(body)) {
            if (pendingSection != null) {
                var paragraph = output.LastChild as Paragraph;
                if (paragraph == null || paragraph.ParagraphProperties?.SectionProperties != null) {
                    paragraph = output.AppendChild(new Paragraph());
                }
                (paragraph.ParagraphProperties ??= new ParagraphProperties()).SectionProperties = pendingSection;
                pendingSection = null;
            }
            if (block is SectionProperties section) pendingSection = section;
            else output.AppendChild(block);
        }
        if (pendingSection != null) output.AppendChild(pendingSection);
        if (!output.HasChildren) output.AppendChild(new Paragraph());
        document.AddMainDocumentPart().Document = new Document(output);
        return projection._omissions.Values.ToArray();
    }

    private void AdmitSource(XElement root, OfficePackageSecurityOptions policy) {
        CheckActive(root, policy.Macros, "macrosPresent", Word + "docSuppData", "macros");
        CheckActive(root, policy.EmbeddedPayloads, "embeddedObjPresent", Word + "docOleData", "embedded OLE content");
        CheckActive(root, policy.ActiveX, "ocxPresent", Word + "ocx", "ActiveX content");
        if (policy.ExternalRelationships == OfficePackageContentPolicy.Reject &&
            root.Descendants(Word + "hlink").Any(element => !string.IsNullOrEmpty((string?)element.Attribute(Word + "dest")))) {
            throw new InvalidDataException("WordML 2003 contains a link rejected by the external relationship policy.");
        }

        void CheckActive(XElement source, OfficePackageContentPolicy contentPolicy, string flag, XName elementName, string label) {
            XAttribute? attribute = source.Attribute(Word + flag);
            if ((attribute == null || !IsTrue(attribute.Value)) && !source.Descendants(elementName).Any()) return;
            if (contentPolicy == OfficePackageContentPolicy.Reject) {
                throw new InvalidDataException("WordML 2003 contains " + label + " rejected by the package content policy.");
            }
            if (attribute != null && IsTrue(attribute.Value)) Omit(attribute.Name, label);
        }
    }

    private IEnumerable<OpenXmlElement> ProjectBlocks(XElement container) {
        foreach (XElement child in container.Elements()) {
            _token.ThrowIfCancellationRequested();
            if (child.Name == Word + "p") {
                yield return ProjectParagraph(child);
            } else if (child.Name == Word + "tbl") {
                yield return ProjectTable(child);
            } else if (child.Name == Auxiliary + "sect" || child.Name == Auxiliary + "sub-section") {
                foreach (OpenXmlElement block in ProjectBlocks(child)) yield return block;
            } else if (child.Name == Word + "sectPr") {
                yield return ProjectSectionProperties(child);
            } else {
                Omit(child, "block content");
            }
        }
    }

    private Paragraph ProjectParagraph(XElement source) {
        var paragraph = new Paragraph();
        foreach (XElement child in source.Elements()) {
            if (child.Name == Word + "pPr") {
                var properties = new ParagraphProperties();
                foreach (XElement property in child.Elements()) {
                    if (property.Name == Word + "sectPr") {
                        properties.SectionProperties = ProjectSectionProperties(property);
                    } else if (property.Name == Word + "jc" && TryJustification((string?)property.Attribute(Word + "val"), out JustificationValues alignment)) {
                        properties.Justification = new Justification { Val = alignment };
                    } else {
                        Omit(property, "paragraph formatting");
                    }
                }
                if (properties.HasChildren) paragraph.ParagraphProperties = properties;
            } else {
                AppendInline(paragraph, child);
            }
        }
        return paragraph;
    }

    private void AppendInline(Paragraph paragraph, XElement source) {
        _token.ThrowIfCancellationRequested();
        if (source.Name == Word + "hlink") {
            Omit(source, "hyperlink behavior; its visible text is retained");
            foreach (XElement child in source.Elements()) AppendInline(paragraph, child);
        } else if (source.Name == Word + "r") {
            var run = new Run();
            foreach (XElement child in source.Elements()) {
                if (child.Name == Word + "t" && !child.HasElements) {
                    run.AppendChild(new Text(child.Value) { Space = SpaceProcessingModeValues.Preserve });
                } else if (child.Name == Word + "tab") {
                    run.AppendChild(new TabChar());
                } else if (child.Name == Word + "br") {
                    var lineBreak = new Break();
                    string? type = (string?)child.Attribute(Word + "type");
                    if (type == "page") lineBreak.Type = BreakValues.Page;
                    else if (type == "column") lineBreak.Type = BreakValues.Column;
                    else if (type != null && type != "text-wrapping") Omit(child, "line break type");
                    string? clear = (string?)child.Attribute(Word + "clear");
                    if (clear == "all") lineBreak.Clear = BreakTextRestartLocationValues.All;
                    else if (clear == "left") lineBreak.Clear = BreakTextRestartLocationValues.Left;
                    else if (clear == "right") lineBreak.Clear = BreakTextRestartLocationValues.Right;
                    else if (clear == "none") lineBreak.Clear = BreakTextRestartLocationValues.None;
                    else if (clear != null) Omit(child, "line break clearing");
                    run.AppendChild(lineBreak);
                } else if (child.Name == Word + "rPr") {
                    run.RunProperties = ProjectRunProperties(child);
                } else {
                    Omit(child, "run content");
                }
            }
            paragraph.AppendChild(run);
        } else {
            Omit(source, "inline content");
        }
    }

    private RunProperties ProjectRunProperties(XElement source) {
        var properties = new RunProperties();
        foreach (XElement child in source.Elements()) {
            string? value = (string?)child.Attribute(Word + "val");
            bool enabled = value == null || IsTrue(value);
            if (child.Name == Word + "b") properties.Bold = new Bold { Val = enabled };
            else if (child.Name == Word + "i") properties.Italic = new Italic { Val = enabled };
            else if (child.Name == Word + "strike") properties.Strike = new Strike { Val = enabled };
            else Omit(child, "character formatting");
        }
        return properties;
    }

    private Table ProjectTable(XElement source) {
        var table = new Table(new TableProperties(), new TableGrid());
        foreach (XElement child in source.Elements()) {
            if (child.Name != Word + "tr") {
                Omit(child, "table formatting or non-row content");
                continue;
            }
            var row = new TableRow();
            foreach (XElement cell in child.Elements()) {
                if (cell.Name != Word + "tc") {
                    Omit(cell, "row formatting or non-cell content");
                    continue;
                }
                var outputCell = new TableCell(ProjectBlocks(cell).Where(block => block is not SectionProperties));
                if (outputCell.LastChild is not Paragraph) outputCell.AppendChild(new Paragraph());
                row.AppendChild(outputCell);
            }
            if (row.HasChildren) table.AppendChild(row);
            else Omit(child, "table row without supported cells");
        }
        if (!table.Elements<TableRow>().Any()) {
            Omit(source, "table without supported cells");
            table.AppendChild(new TableRow(new TableCell(new Paragraph())));
        }
        int columns = table.Elements<TableRow>().Max(row => row.Elements<TableCell>().Count());
        TableGrid grid = table.GetFirstChild<TableGrid>()!;
        for (int i = 0; i < columns; i++) grid.AppendChild(new GridColumn());
        return table;
    }

    private SectionProperties ProjectSectionProperties(XElement source) {
        var properties = new SectionProperties();
        foreach (XElement child in source.Elements()) {
            if (child.Name == Word + "pgSz" &&
                uint.TryParse((string?)child.Attribute(Word + "w"), NumberStyles.None, CultureInfo.InvariantCulture, out uint width) &&
                uint.TryParse((string?)child.Attribute(Word + "h"), NumberStyles.None, CultureInfo.InvariantCulture, out uint height)) {
                var size = new PageSize { Width = width, Height = height };
                string? orientation = (string?)child.Attribute(Word + "orient");
                if (orientation == "landscape") size.Orient = PageOrientationValues.Landscape;
                else if (orientation == "portrait") size.Orient = PageOrientationValues.Portrait;
                else if (orientation != null) Omit(child, "page orientation");
                properties.AppendChild(size);
            } else {
                Omit(child, "section formatting or content");
            }
        }
        return properties;
    }

    private void Omit(XElement element, string reason) => Omit(element.Name, reason);

    private void Omit(XName name, string reason) {
        if (_omissions.ContainsKey(name)) return;
        _omissions.Add(name, new OfficeDiagnostic("WORDML2003_OMITTED", OfficeDiagnosticSeverity.Warning,
            "WordML 2003 " + reason + " was omitted (" + name.LocalName + ").",
            location: name.ToString(), format: "WordML 2003", lossKind: OfficeConversionLossKind.Omission));
    }

    private static bool IsTrue(string value) => value is "1" or "on" or "true" or "yes";

    private static bool TryJustification(string? value, out JustificationValues result) {
        switch (value) {
            case "left": result = JustificationValues.Left; return true;
            case "center": result = JustificationValues.Center; return true;
            case "right": result = JustificationValues.Right; return true;
            case "both": result = JustificationValues.Both; return true;
            default: result = default; return false;
        }
    }
}
