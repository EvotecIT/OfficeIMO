using DocumentFormat.OpenXml.Packaging;
using System.Security.Cryptography;
using System.Text;
using System.Xml;

namespace OfficeIMO.Excel.Xlsb.Write;

/// <summary>Protects retained chart/VBA graphs and detects direct worksheet DOM edits before unchanged-source copying.</summary>
internal sealed class XlsbProjectionPreservationGuard {
    private readonly Dictionary<string, string> _preservedParts;
    private readonly Dictionary<string, string> _worksheets;
    private readonly string _sheetBindings;

    private XlsbProjectionPreservationGuard(ExcelDocument document) {
        _sheetBindings = document.WorkbookRoot.Sheets?.OuterXml ?? string.Empty;
        _preservedParts = CaptureParts(document, worksheets: false);
        _worksheets = CaptureParts(document, worksheets: true);
    }

    internal static XlsbProjectionPreservationGuard Capture(ExcelDocument document) => new(document);

    internal void Validate(ExcelDocument document) {
        if (_sheetBindings != (document.WorkbookRoot.Sheets?.OuterXml ?? string.Empty)
            || !Matches(_preservedParts, CaptureParts(document, worksheets: false)))
            throw new NotSupportedException("Native XLSB rewriting preserves chart sheets, VBA and related parts, but cannot change their content, relationships or sheet bindings. Save those edits as a macro-compatible Open XML workbook.");
    }

    internal bool HasWorksheetChanges(ExcelDocument document) => !Matches(_worksheets, CaptureParts(document, worksheets: true));

    private static bool Matches(Dictionary<string, string> expected, Dictionary<string, string> actual) =>
        expected.Count == actual.Count && expected.All(item => actual.TryGetValue(item.Key, out string? value) && value == item.Value);

    private static Dictionary<string, string> CaptureParts(ExcelDocument document, bool worksheets) {
        var result = new Dictionary<string, string>(StringComparer.Ordinal);
        var visited = new HashSet<OpenXmlPart>(); var queue = new Queue<OpenXmlPart>();
        foreach (var pair in document.OpenXmlDocument.Parts) queue.Enqueue(pair.OpenXmlPart);
        while (queue.Count > 0) {
            OpenXmlPart part = queue.Dequeue();
            if (!visited.Add(part)) continue;
            foreach (var pair in part.Parts) queue.Enqueue(pair.OpenXmlPart);
            bool editablePart = part is WorksheetPart or WorkbookPart or WorkbookStylesPart or SharedStringTablePart;
            if (worksheets ? !editablePart : part is WorksheetPart or WorkbookStylesPart or SharedStringTablePart) continue;
            result.Add(part.Uri.OriginalString, Fingerprint(part, includeContent: worksheets || part is not WorkbookPart));
        }
        return result;
    }

    private static string Fingerprint(OpenXmlPart part, bool includeContent) {
        using var hash = SHA256.Create();
        using (var stream = new CryptoStream(Stream.Null, hash, CryptoStreamMode.Write)) {
            using (var text = new StreamWriter(stream, new UTF8Encoding(false), 4096, leaveOpen: true)) {
                text.Write(part.ContentType); text.Write('\n');
                foreach (var pair in part.Parts.Where(item => item.OpenXmlPart is not SharedStringTablePart)
                    .OrderBy(item => item.RelationshipId, StringComparer.Ordinal))
                    text.Write(pair.RelationshipId + "|" + pair.OpenXmlPart.RelationshipType + "|" + pair.OpenXmlPart.Uri.OriginalString + "\n");
                foreach (var relationship in part.ExternalRelationships.OrderBy(item => item.Id, StringComparer.Ordinal))
                    text.Write(relationship.Id + "|" + relationship.RelationshipType + "|" + relationship.Uri.OriginalString + "\n");
                foreach (var relationship in part.HyperlinkRelationships.OrderBy(item => item.Id, StringComparer.Ordinal))
                    text.Write(relationship.Id + "|hyperlink|" + relationship.Uri.OriginalString + "\n");
                text.Flush();
            }
            if (includeContent) {
                if (part.RootElement is { } root) {
                    using var writer = XmlWriter.Create(stream, new XmlWriterSettings { Encoding = new UTF8Encoding(false), CloseOutput = false, OmitXmlDeclaration = true });
                    root.WriteTo(writer); writer.Flush();
                } else {
                    using var input = part.GetStream(FileMode.Open, FileAccess.Read); input.CopyTo(stream);
                }
            }
            stream.FlushFinalBlock();
        }
        return Convert.ToBase64String(hash.Hash!);
    }
}
