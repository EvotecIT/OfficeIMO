using OfficeIMO.Excel.Xlsb.Model;
using OfficeIMO.Excel.Xlsb.Package;
using OfficeIMO.Excel.Xlsb.Write;

namespace OfficeIMO.Excel.Xlsb;

internal static partial class XlsbWorkbookReader {
    /// <summary>Reads the package-level property parts. They are optional metadata, so a missing or mistyped part is skipped.</summary>
    private static void ReadDocumentPropertyParts(XlsbPackagePartReader parts, XlsbImportOptions options, XlsbWorkbook workbook) {
        IReadOnlyDictionary<string, XlsbPackageRelationship> relationships = parts.ReadRelationships(string.Empty, options.CancellationToken);
        workbook.CorePropertiesXml = ReadPropertyPart(parts, relationships, options,
            XlsbDocumentPropertyParts.CoreRelationshipType, XlsbDocumentPropertyParts.CoreContentType);
        workbook.ExtendedPropertiesXml = ReadPropertyPart(parts, relationships, options,
            XlsbDocumentPropertyParts.ExtendedRelationshipType, XlsbDocumentPropertyParts.ExtendedContentType);
        workbook.CustomPropertiesXml = ReadPropertyPart(parts, relationships, options,
            XlsbDocumentPropertyParts.CustomRelationshipType, XlsbDocumentPropertyParts.CustomContentType);
    }

    private static byte[]? ReadPropertyPart(XlsbPackagePartReader parts,
        IReadOnlyDictionary<string, XlsbPackageRelationship> relationships, XlsbImportOptions options,
        string relationshipType, string contentType) {
        XlsbPackageRelationship? relationship = relationships.Values.FirstOrDefault(item => !item.IsExternal && item.Type == relationshipType);
        if (relationship == null) return null;
        try {
            string partName = XlsbPackagePartReader.ResolveTarget(string.Empty, relationship.Target);
            if (!parts.ContainsPart(partName) || parts.GetContentType(partName) != contentType) return null;
            return parts.ReadPart(partName, options.CancellationToken);
        } catch (Exception exception) when (exception is InvalidDataException or System.Xml.XmlException) {
            // Unreadable optional metadata must not prevent loading the workbook.
            return null;
        }
    }
}
