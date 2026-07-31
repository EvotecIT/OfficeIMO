using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel {
    internal sealed partial class ExcelDocumentReader {
        internal IReadOnlyList<string> GetValidatedWorksheetNames() {
            var names = new List<string>();
            IEnumerable<Sheet> sheets =
                WorkbookRoot.Sheets?.Elements<Sheet>() ?? Enumerable.Empty<Sheet>();
            foreach (Sheet sheet in sheets) {
                _opt.CancellationToken.ThrowIfCancellationRequested();
                string? name = sheet.Name?.Value;
                string? relationshipId = sheet.Id?.Value;
                if (string.IsNullOrEmpty(name) || string.IsNullOrEmpty(relationshipId)) {
                    throw new InvalidDataException(
                        "The OpenXML workbook contains a sheet without a name or relationship id.");
                }
                string validatedName = name!;
                string validatedRelationshipId = relationshipId!;
                if (WorkbookPartRoot.ExternalRelationships.Any(
                    relationship => relationship.Id == validatedRelationshipId)) {
                    throw new InvalidDataException(
                        $"The OpenXML worksheet '{validatedName}' references external relationship '{validatedRelationshipId}'.");
                }

                OpenXmlPart part;
                try {
                    part = WorkbookPartRoot.GetPartById(validatedRelationshipId);
                } catch (Exception exception) when (
                    exception is ArgumentOutOfRangeException
                    || exception is InvalidOperationException
                    || exception is KeyNotFoundException) {
                    throw new InvalidDataException(
                        $"The OpenXML worksheet '{validatedName}' references missing relationship '{validatedRelationshipId}'.",
                        exception);
                }

                if (part is WorksheetPart) {
                    names.Add(validatedName);
                }
            }

            return names;
        }
    }
}
