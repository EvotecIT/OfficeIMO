using OfficeIMO.Excel.OpenXml;
using OfficeIMO.OpenXml.Internal;

namespace OfficeIMO.Excel {
    internal sealed partial class ExcelDocumentReader {
        internal IReadOnlyList<string> GetValidatedWorksheetNames() {
            var names = new List<string>();
            foreach (OfficeXmlNode sheet in WorkbookSheets) {
                _opt.CancellationToken.ThrowIfCancellationRequested();
                string? name = sheet.Attribute("name");
                string? id = WorksheetRelationshipId(sheet);
                if (string.IsNullOrEmpty(name) || string.IsNullOrEmpty(id)) {
                    throw new InvalidDataException("The OpenXML workbook contains a sheet without a name or relationship id.");
                }

                string? partName;
                string? relationshipType;
                bool external;
                if (_liveWorkbook is ExcelPartRef workbook) {
                    ExcelPartRef part = workbook.GetPartById(id);
                    partName = part.Exists ? part.Name : null;
                    relationshipType = part.RelationshipType;
                    external = workbook.GetExternalRelationshipUri(id!) != null;
                } else {
                    var relationship = _package!.FindWorkbookRelationship(id!);
                    partName = relationship.Target;
                    if (partName != null && !_package.PartExists(partName)) partName = null;
                    relationshipType = relationship.Type;
                    external = relationship.External;
                }

                if (external) {
                    throw new InvalidDataException($"The OpenXML worksheet '{name}' references external relationship '{id}'.");
                }
                if (partName == null) {
                    if (ExcelReadWorkbookParts.IsNonWorksheetRelationship(relationshipType)) {
                        throw new InvalidOperationException($"The OpenXML sheet '{name}' references missing relationship '{id}'.");
                    }
                    throw new InvalidDataException($"The OpenXML worksheet '{name}' references missing relationship '{id}'.");
                }
                if (relationshipType == ExcelReadWorkbookParts.WorksheetRelationship) {
                    if (_liveWorkbook == null) _package!.ValidateWorksheetContentType(partName);
                    names.Add(name!);
                } else if (!ExcelReadWorkbookParts.IsNonWorksheetRelationship(relationshipType)) {
                    throw new InvalidDataException($"The OpenXML sheet '{name}' references unsupported internal part '{partName}'.");
                }
            }
            return names;
        }

    }
}
