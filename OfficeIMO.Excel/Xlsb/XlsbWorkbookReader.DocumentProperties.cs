using OfficeIMO.Excel.Xlsb.Model;
using OfficeIMO.Excel.Xlsb.Package;
using OfficeIMO.Excel.Xlsb.Write;

namespace OfficeIMO.Excel.Xlsb;

internal static partial class XlsbWorkbookReader {
    internal const string DocumentPropertiesSkippedDiagnostic = "xlsb-document-properties-skipped";

    /// <summary>
    /// Reads the package-level property parts. They are optional metadata: an unreadable relationship
    /// part or property part is skipped with a diagnostic and never prevents loading the workbook.
    /// </summary>
    private static void ReadDocumentPropertyParts(XlsbPackagePartReader parts, XlsbImportOptions options, XlsbWorkbook workbook) {
        IReadOnlyDictionary<string, XlsbPackageRelationship> relationships;
        try {
            relationships = parts.ReadRelationships(string.Empty, options.CancellationToken);
        } catch (Exception exception) when (exception is InvalidDataException or System.Xml.XmlException) {
            AddDocumentPropertiesSkipped(workbook, "_rels/.rels", "The package relationships could not be read, so document properties were not imported: " + exception.Message);
            return;
        }
        workbook.CorePropertiesXml = ReadPropertyPart(parts, relationships, options, workbook,
            XlsbDocumentPropertyParts.CoreRelationshipType, XlsbDocumentPropertyParts.CoreContentType);
        workbook.ExtendedPropertiesXml = ReadPropertyPart(parts, relationships, options, workbook,
            XlsbDocumentPropertyParts.ExtendedRelationshipType, XlsbDocumentPropertyParts.ExtendedContentType);
        workbook.CustomPropertiesXml = ReadPropertyPart(parts, relationships, options, workbook,
            XlsbDocumentPropertyParts.CustomRelationshipType, XlsbDocumentPropertyParts.CustomContentType);
    }

    internal static void AddDocumentPropertiesSkipped(XlsbWorkbook workbook, string? partName, string message) =>
        workbook.AddDiagnostic(new XlsbImportDiagnostic(DocumentPropertiesSkippedDiagnostic, XlsbImportDiagnosticSeverity.Warning, message, partName));

    private static byte[]? ReadPropertyPart(XlsbPackagePartReader parts,
        IReadOnlyDictionary<string, XlsbPackageRelationship> relationships, XlsbImportOptions options,
        XlsbWorkbook workbook, string relationshipType, string contentType) {
        XlsbPackageRelationship? relationship = relationships.Values.FirstOrDefault(item => !item.IsExternal && item.Type == relationshipType);
        if (relationship == null) return null;
        string? partName = null;
        try {
            partName = XlsbPackagePartReader.ResolveTarget(string.Empty, relationship.Target);
            if (!parts.ContainsPart(partName)) {
                AddDocumentPropertiesSkipped(workbook, partName, "The related document-property part is missing.");
                return null;
            }
            string actualType = parts.GetContentType(partName);
            if (actualType != contentType) {
                AddDocumentPropertiesSkipped(workbook, partName, $"The document-property part has content type '{actualType}' instead of '{contentType}'.");
                return null;
            }
            return parts.ReadPart(partName, options.CancellationToken);
        } catch (Exception exception) when (exception is InvalidDataException or System.Xml.XmlException) {
            AddDocumentPropertiesSkipped(workbook, partName, "The document-property part could not be read: " + exception.Message);
            return null;
        }
    }
}
