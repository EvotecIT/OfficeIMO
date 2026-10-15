using OfficeIMO.Excel.Xlsb.Package;

namespace OfficeIMO.Excel.Xlsb.Read {
    internal sealed partial class XlsbTabularWorkbook {
        private const string SharedStringsRelationshipSuffix = "/sharedStrings";

        private IReadOnlyList<string> ReadSharedStrings(
            string workbookPartName,
            IReadOnlyDictionary<string, XlsbPackageRelationship> relationships) {
            XlsbPackageRelationship? relationship = XlsbPackagePartReader.GetOptionalSingletonRelationship(
                relationships, SharedStringsRelationshipSuffix, "shared-string");
            if (relationship == null) return Array.Empty<string>();
            string partName = XlsbPackagePartReader.ResolveTarget(workbookPartName, relationship.Target);
            return XlsbSharedStringTable.Read(
                _parts.ReadSeekablePart(partName, _cancellationToken), partName, _limits,
                _recordBudget, _maxSharedStringItemCharacters, _maxSharedStringCharacters,
                CheckCancellation);
        }
    }
}
