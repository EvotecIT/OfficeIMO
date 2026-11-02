using System.Xml;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using static OfficeIMO.Excel.SpreadsheetXml.SpreadsheetXmlContext;

namespace OfficeIMO.Excel.SpreadsheetXml;

internal static class SpreadsheetXmlMetadata {
    internal static void ReadCell(XElement source, ExcelSheet sheet, int row, int column, SpreadsheetXmlContext context) {
        string reference = ExcelA1.CellReference(row, column);
        if ((string?)source.Attribute(Ss + "HRef") is string link) {
            string? tooltip = (string?)source.Attribute(X + "HRefScreenTip");
            if (link.StartsWith("#", StringComparison.Ordinal)) sheet.AddInternalHyperlinkReference(reference, link.Substring(1), style: false, normalizeLocation: false, tooltip: tooltip);
            else sheet.SetHyperlinkReference(row, column, link, style: false, tooltip: tooltip);
        }
        XElement[] comments = source.Elements(Ss + "Comment").ToArray();
        if (comments.Length > 1) throw new InvalidDataException("Multiple comments on an XML Spreadsheet cell.");
        if (comments.Length == 0) return;
        XElement comment = comments[0];
        context.Inspect(comment, new[] { "Data" }, new[] { "Author", "ShowAlways" });
        XElement? data = comment.Element(Ss + "Data");
        string text = data?.Value ?? string.Empty;
        if (text.Length > 32767) throw new InvalidDataException("XML Spreadsheet comment exceeds the text limit.");

        if ((string?)comment.Attribute(Ss + "ShowAlways") == "1") context.Omitted("comment visibility");
        if (text.Length > 0) {
            string author = (string?)comment.Attribute(Ss + "Author") ?? string.Empty;
            if (data?.HasElements == true) sheet.SetCommentRichText(row, column, SpreadsheetXmlRichText.Read(data, context), author);
            else sheet.SetComment(row, column, text, author);
        }
    }

    internal static void ReadNames(XElement? source, WorkbookPart workbook, uint? scope, SpreadsheetXmlContext context) {
        if (source == null) return;
        context.Inspect(source, new[] { "NamedRange" }, Array.Empty<string>());
        DefinedNames names = workbook.Workbook!.DefinedNames ??= new DefinedNames();
        foreach (XElement element in source.Elements(Ss + "NamedRange")) {
            context.Options.CancellationToken.ThrowIfCancellationRequested();
            context.Inspect(element, Array.Empty<string>(), new[] { "Name", "RefersTo", "Hidden" });
            string name = (string?)element.Attribute(Ss + "Name") ?? throw new InvalidDataException("Missing defined name.");
            name = name is "Print_Area" or "Print_Titles" or "_FilterDatabase" ? "_xlnm." + name : ExcelDocument.NormalizeDefinedName(name, ExcelDefinedNameValidationMode.Strict);
            context.DefinedName(name, scope);
            string formula = (string?)element.Attribute(Ss + "RefersTo") ?? throw new InvalidDataException("Missing defined-name reference.");
            if (formula.Length > 8193 || !formula.StartsWith("=", StringComparison.Ordinal)) throw new InvalidDataException("Invalid defined-name formula.");
            names.Append(new DefinedName {
                Name = name, LocalSheetId = scope, Hidden = (string?)element.Attribute(Ss + "Hidden") is string hidden && SpreadsheetXmlGeometry.Flag(hidden),
                Text = ExcelFormulaSyntaxTree.Parse(formula.Substring(1), ExcelReferenceStyle.R1C1, 1, 1).ConvertReferences(ExcelReferenceStyle.A1)
            });
        }
    }

    internal static void WriteNames(XmlWriter writer, WorkbookPart workbook, uint? scope, SpreadsheetXmlContext context) {
        var names = workbook.Workbook!.DefinedNames?.Elements<DefinedName>().Where(item => item.LocalSheetId?.Value == scope).ToArray() ?? Array.Empty<DefinedName>();
        if (names.Length == 0) return;
        writer.WriteStartElement("Names", Ss.NamespaceName);
        foreach (DefinedName name in names) {
            context.Options.CancellationToken.ThrowIfCancellationRequested();
            writer.WriteStartElement("NamedRange", Ss.NamespaceName);
            string text = name.Name?.Value ?? throw new InvalidDataException("Unnamed defined name.");
            if (text.StartsWith("_xlnm.", StringComparison.Ordinal)) text = text.Substring(6);
            writer.WriteAttributeString("ss", "Name", Ss.NamespaceName, text);
            writer.WriteAttributeString("ss", "RefersTo", Ss.NamespaceName, "=" + ExcelFormulaSyntaxTree.Parse(name.Text).ConvertReferences(ExcelReferenceStyle.R1C1, 1, 1));
            if (name.Hidden?.Value == true) writer.WriteAttributeString("ss", "Hidden", Ss.NamespaceName, "1");
            foreach (var attribute in name.GetAttributes())
                if (attribute.LocalName is not "name" and not "localSheetId" and not "hidden") context.Omitted("defined-name metadata " + attribute.LocalName);
            writer.WriteEndElement();
        }
        writer.WriteEndElement();
    }

    internal static Dictionary<string, XElement> WriteCellMetadata(WorksheetPart part, SpreadsheetXmlContext context) {
        var result = new Dictionary<string, XElement>(StringComparer.OrdinalIgnoreCase);
        foreach (Hyperlink link in part.Worksheet!.GetFirstChild<Hyperlinks>()?.Elements<Hyperlink>() ?? Enumerable.Empty<Hyperlink>()) {
            string reference = link.Reference?.Value ?? string.Empty;
            if (!ExcelReference.TryParse(reference, out var parsed) || parsed?.Kind != ExcelReferenceKind.Cell) { context.Omitted("range hyperlinks"); continue; }
            string? target = link.Id != null ? part.HyperlinkRelationships.FirstOrDefault(item => item.Id == link.Id.Value)?.Uri.OriginalString : null;
            if (link.Location?.Value is string location) target = target == null ? "#" + location : target + "#" + location;
            if (target == null) { context.Omitted("unresolved hyperlink"); continue; }
            XElement cell = Get(reference);
            cell.SetAttributeValue(Ss + "HRef", target);
            if (link.Tooltip?.Value is string tooltip) cell.SetAttributeValue(X + "HRefScreenTip", tooltip);
        }
        var comments = part.WorksheetCommentsPart?.Comments;
        string[] authors = comments?.Authors?.Elements<Author>().Select(item => item.Text).ToArray() ?? Array.Empty<string>();
        foreach (Comment comment in comments?.CommentList?.Elements<Comment>() ?? Enumerable.Empty<Comment>()) {
            string reference = comment.Reference?.Value ?? throw new InvalidDataException("Comment has no cell reference.");
            uint author = comment.AuthorId?.Value ?? 0;
            if (author >= authors.Length) throw new InvalidDataException("Invalid comment author index.");
            var node = new XElement(Ss + "Comment", new XAttribute(Ss + "Author", authors[author]), new XElement(Ss + "Data", comment.CommentText?.InnerText ?? string.Empty));
            if (comment.CommentText?.Elements<Run>().Any() == true) {
                XElement data = node.Element(Ss + "Data")!; data.RemoveNodes();
                using var writer = data.CreateWriter();
                SpreadsheetXmlRichText.Write(writer, comment.CommentText, part.GetParentParts().OfType<WorkbookPart>().Single(), context);
            }
            Get(reference).Add(node);
        }
        return result;
        XElement Get(string reference) {
            if (!result.TryGetValue(reference, out XElement? item)) result[reference] = item = new XElement(Ss + "Cell");
            return item;
        }
    }
}
