using DocumentFormat.OpenXml.Packaging;

namespace OfficeIMO.Excel;

/// <summary>
/// Identifies a package part created or updated by an Excel metadata operation.
/// </summary>
public sealed class ExcelPackagePartInfo {
    internal ExcelPackagePartInfo(OpenXmlPartContainer owner, OpenXmlPart part) {
        RelationshipId = owner.GetIdOfPart(part);
        Uri = part.Uri.ToString();
        ContentType = part.ContentType;
        RelationshipType = part.RelationshipType;
    }

    /// <summary>Gets the relationship identifier used by the owning workbook or worksheet part.</summary>
    public string RelationshipId { get; }

    /// <summary>Gets the package-relative URI of the part.</summary>
    public string Uri { get; }

    /// <summary>Gets the MIME content type of the part.</summary>
    public string ContentType { get; }

    /// <summary>Gets the relationship type that describes the part's role.</summary>
    public string RelationshipType { get; }
}
