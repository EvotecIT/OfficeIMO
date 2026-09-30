using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        private RichValueErrorLookup? _richValueErrorLookup;
        private object? _richValueErrorMetadata;
        private object? _richValueErrorValues;
        private object? _richValueErrorStructures;
        private int _richValueErrorMetadataCount, _richValueErrorValueCount, _richValueErrorStructureCount;

        private string? ResolveRichValueError(Cell cell, string? fallback) {
            if (fallback == null || cell.DataType?.Value != DocumentFormat.OpenXml.Spreadsheet.CellValues.Error || cell.ValueMetaIndex == null) return fallback;
            var workbook = _excelDocument.WorkbookPartRoot;
            // Editable documents expose their Open XML roots. A direct edit can
            // change a rich value without changing any collection count.
            if (_excelDocument.FileOpenAccess != FileAccess.Read && _formulaEvaluationCache == null)
                return RichValueErrorLookup.FromWorkbook(workbook).Resolve(cell.ValueMetaIndex.Value, fallback);
            // Validate raw parts before touching properties that materialize SDK roots.
            bool initialized = _richValueErrorLookup == null;
            if (initialized) _richValueErrorLookup = RichValueErrorLookup.FromWorkbook(workbook);
            var metadata = workbook.CellMetadataPart?.Metadata;
            var values = workbook.RdRichValueParts.FirstOrDefault()?.RichValueData;
            var structures = workbook.GetPartsOfType<DocumentFormat.OpenXml.Packaging.RdRichValueStructurePart>().FirstOrDefault()?.RichValueStructures;
            int metadataCount = metadata?.GetFirstChild<ValueMetadata>()?.ChildElements.Count ?? 0;
            int valueCount = values?.ChildElements.Count ?? 0;
            int structureCount = structures?.ChildElements.Count ?? 0;
            if (initialized || !ReferenceEquals(_richValueErrorMetadata, metadata)
                || !ReferenceEquals(_richValueErrorValues, values) || !ReferenceEquals(_richValueErrorStructures, structures)
                || _richValueErrorMetadataCount != metadataCount || _richValueErrorValueCount != valueCount || _richValueErrorStructureCount != structureCount) {
                if (!initialized) _richValueErrorLookup = RichValueErrorLookup.FromWorkbook(workbook);
                _richValueErrorMetadata = metadata; _richValueErrorValues = values; _richValueErrorStructures = structures;
                _richValueErrorMetadataCount = metadataCount; _richValueErrorValueCount = valueCount; _richValueErrorStructureCount = structureCount;
            }
            return _richValueErrorLookup!.Resolve(cell.ValueMetaIndex.Value, fallback);
        }
    }
}
