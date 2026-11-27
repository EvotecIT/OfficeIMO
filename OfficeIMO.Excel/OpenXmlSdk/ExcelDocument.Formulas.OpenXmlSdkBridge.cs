using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel {
    /// <summary>The formula model of a workbook the SDK package holds. It goes with the sheet's SDK formula model.</summary>
    public partial class ExcelDocument {
        private SdkFormulaWorkbook? _sdkFormulaWorkbook;

        private ExcelFormulaWorkbookModel SdkFormulaWorkbookModel => _sdkFormulaWorkbook ??= new SdkFormulaWorkbook(this);

        private sealed class SdkFormulaWorkbook : ExcelFormulaWorkbookModel {
            private readonly ExcelDocument _document;

            internal SdkFormulaWorkbook(ExcelDocument document) => _document = document;

            private IEnumerable<(Sheet Sheet, WorksheetPart? Part)> Sheets() {
                WorkbookPart workbook = _document.WorkbookPartRoot;
                foreach (Sheet sheet in _document.WorkbookRoot.Sheets?.Elements<Sheet>() ?? Enumerable.Empty<Sheet>()) {
                    string? relationshipId = sheet.Id?.Value;
                    WorksheetPart? part = null;
                    if (!string.IsNullOrWhiteSpace(relationshipId)) {
                        foreach (IdPartPair pair in workbook.Parts) {
                            if (pair.OpenXmlPart is WorksheetPart worksheet && string.Equals(pair.RelationshipId, relationshipId, StringComparison.Ordinal)) {
                                part = worksheet;
                                break;
                            }
                        }
                    }
                    yield return (sheet, part);
                }
            }

            internal override List<ExcelFormulaSheetEntry> SheetEntries() =>
                Sheets().Select(item => new ExcelFormulaSheetEntry(item.Sheet.Name?.Value, item.Part != null, item.Sheet)).ToList();

            internal override List<ExcelFormulaDefinedName> DefinedNames() {
                var names = new List<ExcelFormulaDefinedName>();
                foreach (DefinedName definedName in _document.WorkbookRoot.DefinedNames?.Elements<DefinedName>() ?? Enumerable.Empty<DefinedName>())
                    names.Add(new ExcelFormulaDefinedName(definedName.Name?.Value, definedName.Text, definedName.LocalSheetId?.Value));
                return names;
            }

            // An entry read from the workbook's XML before the SDK package opened is found again by its name.
            internal override ExcelSheet CreateSheet(ExcelFormulaSheetEntry entry) => entry.Element is Sheet sheet
                ? new ExcelSheet(_document, _document._spreadSheetDocument, sheet)
                : _document.FindSheetForFormulas(entry.Name ?? string.Empty) ?? throw new InvalidOperationException("The worksheet '" + entry.Name + "' is no longer part of the workbook.");

            internal override IEnumerable<(ExcelFormulaSheetEntry Sheet, ExcelFormulaTable Table)> Tables() {
                foreach ((Sheet sheet, WorksheetPart? part) in Sheets()) {
                    if (part == null) continue;
                    var entry = new ExcelFormulaSheetEntry(sheet.Name?.Value, true, sheet);
                    foreach (Table table in part.TableDefinitionParts.Select(tablePart => tablePart.Table).OfType<Table>())
                        yield return (entry, ExcelSheet.ToFormulaTableForWorkbook(table));
                }
            }

            internal override bool HasCellMetadata => _document.WorkbookPartRoot.CellMetadataPart != null;

            internal override ExcelCellMetadataIndex? CellMetadata {
                get {
                    CellMetadataPart? part = _document.WorkbookPartRoot.CellMetadataPart;
                    if (part == null) return null;
                    if (!part.IsRootElementLoaded) ExcelSheet.ValidateCellMetadataPart(part);
                    if (part.Metadata is not Metadata metadata) return null;
                    byte[] xml = System.Text.Encoding.UTF8.GetBytes(metadata.OuterXml);
                    return ExcelCellMetadataIndex.Read(xml, xml.Length);
                }
            }

            internal override bool RequestsFullCalculation(ExcelSheet sheet) {
                CalculationProperties? calculation = _document.WorkbookRoot.GetFirstChild<CalculationProperties>();
                return calculation?.FullCalculationOnLoad?.Value == true
                    || calculation?.ForceFullCalculation?.Value == true
                    || sheet.SheetRequestsFullCalculationOnLoad;
            }
        }
    }
}
