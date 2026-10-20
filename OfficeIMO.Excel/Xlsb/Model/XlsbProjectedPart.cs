using OfficeIMO.Excel.Xlsb.Package;

namespace OfficeIMO.Excel.Xlsb.Model;

/// <summary>A bounded original XML or binary part linked from a projected chart or VBA carrier.</summary>
internal sealed class XlsbProjectedPart {
    internal XlsbProjectedPart(string contentType, byte[] bytes, IReadOnlyDictionary<string, XlsbPackageRelationship> relationships) {
        ContentType = contentType; Bytes = bytes; Relationships = relationships;
    }
    internal string ContentType { get; }
    internal byte[] Bytes { get; }
    internal IReadOnlyDictionary<string, XlsbPackageRelationship> Relationships { get; }
}
