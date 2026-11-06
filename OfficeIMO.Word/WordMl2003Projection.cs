using System.Threading;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word;

/// <summary>Projects the supported Word 2003 XML vocabulary into ordinary editable Word parts.</summary>
internal sealed partial class WordMl2003Projection {
    private static readonly XNamespace Word = WordXmlInterchange.WordMl;
    private static readonly XNamespace Auxiliary = "http://schemas.microsoft.com/office/word/2003/auxHint";
    private static readonly XNamespace ModernWord = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private readonly CancellationToken _token;
    private readonly Dictionary<string, OfficeDiagnostic> _omissions = new(StringComparer.Ordinal);
    private MainDocumentPart _main = null!;
    private OpenXmlPart _story = null!;

    private WordMl2003Projection(CancellationToken token) => _token = token;

    internal static IReadOnlyList<OfficeDiagnostic> WritePackage(XElement root, Stream destination,
        OfficePackageSecurityOptions policy, CancellationToken token) {
        var projection = new WordMl2003Projection(token) { _resourcePolicy = policy };
        XElement? body = root.Element(Word + "body");
        if (body == null || body.ElementsAfterSelf(Word + "body").Any())
            throw new InvalidDataException("WordML 2003 must contain exactly one Word body.");
        projection.ValidateDocumentSingletons(root);
        projection.InitializeReconstructionLimits(root, policy);
        projection.AdmitSource(root, policy);
        using var document = WordprocessingDocument.Create(destination, WordprocessingDocumentType.Document);
        projection.ReserveRelationship(document, createsPart: true);
        projection._main = document.AddMainDocumentPart();
        projection._story = projection._main;
        projection.ReadDocumentParts(root);
        var output = new Body();
        SectionProperties? pendingSection = null;
        foreach (OpenXmlElement block in projection.ProjectBlocks(body)) {
            if (pendingSection != null) {
                var paragraph = output.LastChild as Paragraph;
                if (paragraph == null || paragraph.ParagraphProperties?.SectionProperties != null)
                    paragraph = output.AppendChild(new Paragraph());
                (paragraph.ParagraphProperties ??= new ParagraphProperties()).SectionProperties = pendingSection;
                pendingSection = null;
            }
            if (block is SectionProperties section) pendingSection = section;
            else output.AppendChild(block);
        }
        if (pendingSection != null) output.AppendChild(pendingSection);
        if (!output.HasChildren) output.AppendChild(new Paragraph());
        projection._main.Document = new Document(output);
        return projection._omissions.Values.ToArray();
    }

    private void AdmitSource(XElement root, OfficePackageSecurityOptions policy) {
        CheckActive(policy.Macros, "macrosPresent", Word + "docSuppData", "macros");
        CheckActive(policy.EmbeddedPayloads, "embeddedObjPresent", Word + "docOleData", "embedded OLE content");
        CheckActive(policy.ActiveX, "ocxPresent", Word + "ocx", "ActiveX content");
        if (policy.ExternalRelationships == OfficePackageContentPolicy.Reject &&
            root.Descendants(Word + "hlink").Any(element => !string.IsNullOrEmpty((string?)element.Attribute(Word + "dest"))))
            throw new InvalidDataException("WordML 2003 contains a link rejected by the external relationship policy.");

        void CheckActive(OfficePackageContentPolicy contentPolicy, string flag, XName elementName, string label) {
            XAttribute? attribute = root.Attribute(Word + flag);
            if ((attribute == null || !IsTrue(attribute.Value)) && !root.Descendants(elementName).Any()) return;
            if (contentPolicy == OfficePackageContentPolicy.Reject)
                throw new InvalidDataException("WordML 2003 contains " + label + " rejected by the package content policy.");
            if (attribute != null && IsTrue(attribute.Value)) Omit(attribute.Name, label);
        }
    }

    private IEnumerable<OpenXmlElement> ProjectBlocks(XElement container, bool inTableCell = false) {
        CheckAttributes(container, container.Name == Word + "hdr" || container.Name == Word + "ftr" ? "type" : "");
        foreach (XElement child in container.Elements()) {
            _token.ThrowIfCancellationRequested();
            if (child.Name == Word + "p") yield return ProjectParagraph(child, inTableCell);
            else if (child.Name == Word + "tbl") yield return ProjectTable(child);
            else if (child.Name == Auxiliary + "sect" || child.Name == Auxiliary + "sub-section") {
                foreach (OpenXmlElement block in ProjectBlocks(child, inTableCell)) yield return block;
            } else if (child.Name == Word + "sectPr") {
                if (inTableCell || _story != _main)
                    Omit(child, inTableCell ? "section properties inside a table cell" : "section properties inside a header or footer");
                else yield return ProjectSectionProperties(child);
            } else if (inTableCell && child.Name == Word + "tcPr") {
                // Cell properties are projected by the containing table.
            } else Omit(child, "block content");
        }
    }

    private Paragraph ProjectParagraph(XElement source, bool inTableCell) {
        CheckAttributes(source, "");
        var paragraph = new Paragraph();
        foreach (XElement child in source.Elements()) {
            if (child.Name == Word + "pPr") {
                XElement mapped = MapProperties(child, "pPr")!;
                XElement? section = child.Element(Word + "sectPr");
                var properties = new ParagraphProperties(mapped.ToString(SaveOptions.DisableFormatting));
                if (section != null) {
                    if (inTableCell || _story != _main)
                        Omit(section, inTableCell ? "section properties inside a table cell" : "section properties inside a header or footer");
                    else properties.SectionProperties = ProjectSectionProperties(section);
                }
                if (properties.HasChildren) paragraph.ParagraphProperties = properties;
            } else AppendInline(paragraph, child);
        }
        return paragraph;
    }

    private void AppendInline(OpenXmlCompositeElement target, XElement source) {
        _token.ThrowIfCancellationRequested();
        if (source.Name == Word + "hlink") {
            AppendHyperlink(target, source);
        } else if (source.Name == Word + "r") {
            CheckAttributes(source, "");
            var run = new Run();
            foreach (XElement child in source.Elements()) {
                if (child.Name == Word + "t" && !child.HasElements) {
                    CheckAttributes(child, "");
                    run.AppendChild(new Text(child.Value) { Space = SpaceProcessingModeValues.Preserve });
                } else if (child.Name == Word + "tab") {
                    CheckAttributes(child, "");
                    run.AppendChild(new TabChar());
                } else if (child.Name == Word + "br") {
                    XElement? mapped = MapProperties(child, "br");
                    if (mapped == null) continue;
                    var parsed = new Run(new XElement(ModernWord + "r", mapped).ToString(SaveOptions.DisableFormatting));
                    run.AppendChild(parsed.GetFirstChild<Break>()!.CloneNode(true));
                } else if (child.Name == Word + "rPr") {
                    run.RunProperties = new RunProperties(MapProperties(child, "rPr")!.ToString(SaveOptions.DisableFormatting));
                } else if (child.Name == Word + "pict") {
                    Picture? picture = ProjectPicture(child);
                    if (picture != null) run.AppendChild(picture);
                } else if (child.Name == Word + "instrText" && !child.HasElements) {
                    CheckAttributes(child, "");
                    run.AppendChild(new FieldCode(child.Value) { Space = SpaceProcessingModeValues.Preserve });
                } else if (child.Name == Word + "fldChar") {
                    XElement? mapped = MapProperties(child, "fldChar");
                    if (mapped != null) run.AppendChild(new FieldChar(mapped.ToString(SaveOptions.DisableFormatting)));
                } else Omit(child, "run content");
            }
            target.AppendChild(run);
        } else Omit(source, "inline content");
    }

    private Table ProjectTable(XElement source) {
        CheckAttributes(source, "");
        XElement? sourceGrid = source.Element(Word + "tblGrid");
        int columns = AdmitTableGrid(source, sourceGrid);
        var table = new Table();
        XElement? properties = source.Element(Word + "tblPr");
        table.AppendChild(properties == null ? new TableProperties() :
            new TableProperties(MapProperties(properties, "tblPr")!.ToString(SaveOptions.DisableFormatting)));
        var grid = sourceGrid == null ? new TableGrid() :
            new TableGrid(MapProperties(sourceGrid, "tblGrid")!.ToString(SaveOptions.DisableFormatting));
        table.AppendChild(grid);
        foreach (XElement child in source.Elements()) {
            if (child.Name == Word + "tblPr" || child.Name == Word + "tblGrid") continue;
            if (child.Name != Word + "tr") { Omit(child, "table content"); continue; }
            CheckAttributes(child, "");
            var row = new TableRow();
            foreach (XElement cell in child.Elements()) {
                if (cell.Name == Word + "trPr") {
                    row.AppendChild(new TableRowProperties(MapProperties(cell, "trPr")!.ToString(SaveOptions.DisableFormatting)));
                } else if (cell.Name == Word + "tblPrEx") {
                    row.AppendChild(new TablePropertyExceptions(MapProperties(cell, "tblPrEx")!.ToString(SaveOptions.DisableFormatting)));
                } else if (cell.Name == Word + "tc") {
                    var outputCell = new TableCell();
                    XElement? cellProperties = cell.Element(Word + "tcPr");
                    if (cellProperties != null)
                        outputCell.AppendChild(new TableCellProperties(MapProperties(cellProperties, "tcPr")!.ToString(SaveOptions.DisableFormatting)));
                    if (outputCell.TableCellProperties?.GridSpan?.Val?.Value == 0)
                        outputCell.TableCellProperties.GridSpan.Val = 1;
                    outputCell.Append(ProjectBlocks(cell, inTableCell: true));
                    if (outputCell.LastChild is not Paragraph) outputCell.AppendChild(new Paragraph());
                    row.AppendChild(outputCell);
                } else Omit(cell, "row content");
            }
            if (row.Elements<TableCell>().Any()) table.AppendChild(row);
            else Omit(child, "table row without supported cells");
        }
        if (!table.Elements<TableRow>().Any()) {
            Omit(source, "table without supported cells");
            table.AppendChild(new TableRow(new TableCell(new Paragraph())));
        }
        if (!grid.HasChildren) {
            for (int i = 0; i < columns; i++) grid.AppendChild(new GridColumn());
        }
        return table;
    }

    private SectionProperties ProjectSectionProperties(XElement source) {
        var stories = new HashSet<(XName Name, string? Type)>();
        foreach (XElement child in source.Elements().Where(element => element.Name == Word + "hdr" || element.Name == Word + "ftr")) {
            if (!stories.Add((child.Name, (string?)child.Attribute(Word + "type"))))
                throw new InvalidDataException("WordML 2003 contains a duplicate header or footer type in a section.");
        }
        var properties = new SectionProperties(MapProperties(source, "sectPr")!.ToString(SaveOptions.DisableFormatting));
        foreach (XElement child in source.Elements().Where(element => element.Name == Word + "hdr" || element.Name == Word + "ftr"))
            AppendHeaderFooter(properties, child);
        return properties;
    }

    private void Omit(XElement element, string reason) => Omit(element.Name, reason);

    private void Omit(XName name, string reason) {
        string identity = name + "|" + reason;
        if (_omissions.ContainsKey(identity)) return;
        _omissions.Add(identity, new OfficeDiagnostic("WORDML2003_OMITTED", OfficeDiagnosticSeverity.Warning,
            "WordML 2003 " + reason + " was omitted (" + name.LocalName + ").",
            location: name.ToString(), format: "WordML 2003", lossKind: OfficeConversionLossKind.Omission));
    }

    private static bool IsTrue(string value) => value is "1" or "on" or "true" or "yes";
}
